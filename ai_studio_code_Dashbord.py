import streamlit as st
import pandas as pd
import io
from rapidfuzz import process
import xlsxwriter

st.set_page_config(page_title="المحلل الاحترافي الذكي", layout="wide")

# --- دالة توحيد المسميات الذكية ---
def unify_names(series):
    series = series.fillna("غير محدد").astype(str)
    counts = series.value_counts()
    unique_names = counts.index.tolist()
    mapping = {}
    processed = set()
    for name in unique_names:
        if name in processed: continue
        matches = process.extract(name, unique_names, score_cutoff=85)
        group = [m[0] for m in matches]
        max_freq = counts[group].max()
        best_matches = [n for n in group if counts[n] == max_freq]
        best_match = min(best_matches, key=len) # الأقصر طولاً عند التساوي
        for m in group:
            mapping[m] = best_match
            processed.add(m)
    return series.map(mapping)

# --- دالة إنشاء التقرير النصي المفصل ---
def generate_long_report(df, selected_cols):
    report_lines = ["تقرير تحليل البيانات التنفيذي", "====================", ""]
    for col in selected_cols:
        report_lines.append(f"تحليل محور: {col}")
        if df[col].dtype in ['int64', 'float64']:
            desc = df[col].describe()
            report_lines.append(f"- تشير القراءات الإحصائية لعمود ({col}) إلى أن المتوسط العام بلغ {desc['mean']:.2f}، مع تسجيل أعلى قيمة عند {desc['max']} وأدنى قيمة عند {desc['min']}.")
            report_lines.append(f"- يتضح من توزيع البيانات أن هناك تبايناً قدره {desc['std']:.2f}، مما يستدعي من أصحاب المصلحة مراجعة القيم التي تبتعد عن المتوسط لضمان جودة العمليات.")
        else:
            top_val = df[col].mode()[0]
            count_top = (df[col] == top_val).sum()
            total = len(df)
            perc = (count_top / total) * 100
            report_lines.append(f"- بعد معالجة البيانات وتوحيد المسميات، تبين أن الفئة الأكثر تكراراً هي '{top_val}'، حيث ظهرت في {count_top} سجلات بنسبة {perc:.1f}% من إجمالي البيانات.")
            report_lines.append(f"- هذا التركز في فئة معينة يتطلب توجيه الموارد لدعم هذا القطاع أو دراسة أسباب التكرار العالي فيه مقارنة بالفئات الأخرى.")
        report_lines.append("")
    return "\n".join(report_lines)

# --- الواجهة ---
st.title("🚀 نظام تحليل البيانات وإعداد التقارير التنفيذية")

uploaded_file = st.file_uploader("ارفع ملف الإكسل", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("اختر ورقة العمل:", xl.sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    
    selected_cols = st.multiselect("اختر الأعمدة للتحليل:", df.columns.tolist(), default=df.columns.tolist())
    
    if st.button("توليد التقرير والرسوم البيانية القابلة للتعديل"):
        with st.spinner('جاري التحليل العميق وتوليد ملف الإكسل...'):
            # 1. المعالجة
            df_processed = df[selected_cols].copy()
            for col in df_processed.select_dtypes(include=['object']):
                df_processed[col] = unify_names(df_processed[col])
            
            # 2. توليد التقرير النصي
            full_report = generate_long_report(df_processed, selected_cols)
            
            # 3. إنشاء ملف Excel باستخدام XlsxWriter لإنشاء رسوم أصلية
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            
            # أ. صفحة البيانات
            df_processed.to_excel(writer, sheet_name='البيانات المعالجة', index=False)
            
            # ب. صفحة التقرير النصي
            report_df = pd.DataFrame({'التقرير التنفيذي': full_report.split('\n')})
            report_df.to_excel(writer, sheet_name='التقرير التحليلي', index=False)
            
            # ج. صفحة الرسوم البيانية (Dashboard)
            workbook = writer.book
            dashboard_sheet = workbook.add_worksheet('لوحة المعلومات')
            
            row_cursor = 1
            for col in selected_cols:
                # تجهيز بيانات الرسم البياني
                summary_sheet_name = f"Summary_{col[:20]}" # اسم ورقة العمل للبيانات المختصرة
                if df_processed[col].dtype in ['int64', 'float64']:
                    summary = df_processed[col].value_counts(bins=5).sort_index()
                    chart_type = 'column' # أعمدة للبيانات الرقمية
                else:
                    summary = df_processed[col].value_counts().head(10)
                    chart_type = 'pie' if len(summary) <= 5 else 'bar' # دائرة إذا كانت الفئات قليلة
                
                summary.to_excel(writer, sheet_name=summary_sheet_name)
                
                # إنشاء الرسم البياني في الإكسل
                chart = workbook.add_chart({'type': chart_type})
                chart.add_series({
                    'name':       f'تحليل {col}',
                    'categories': f"='{summary_sheet_name}'!$A$2:$A${len(summary)+1}",
                    'values':     f"='{summary_sheet_name}'!$B$2:$B${len(summary)+1}",
                })
                chart.set_title({'name': f'توزيع بيانات {col}'})
                chart.set_style(10)
                
                dashboard_sheet.insert_chart(f'B{row_cursor}', chart)
                row_cursor += 20 # ترك مسافة بين الرسوم

            writer.close()
            
            st.success("تم تجهيز التقرير بنجاح!")
            st.text_area("معاينة التقرير التنفيذي:", full_report, height=300)
            
            st.download_button(
                label="📥 تحميل ملف الإكسل النهائي (البيانات + الرسوم + التقرير)",
                data=output.getvalue(),
                file_name="Executive_Analysis_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
