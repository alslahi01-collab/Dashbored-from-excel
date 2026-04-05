import streamlit as st
import pandas as pd
import io
import re
from rapidfuzz import process
import xlsxwriter

st.set_page_config(page_title="المحلل الاحترافي الذكي", layout="wide")

# --- دالة لتنظيف أسماء الصفحات لتتوافق مع شروط إكسل ---
def clean_sheet_name(name, index):
    # حذف الرموز الممنوعة: [ ] : * ? / \
    clean_name = re.sub(r'[\[\]:*?/\\]', '', str(name))
    # اختصار الاسم لـ 25 حرفاً لترك مساحة للبادئة
    clean_name = clean_name[:25]
    # إضافة رقم الفهرس لضمان عدم تكرار الأسماء
    return f"Data_{index}_{clean_name}"

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
        best_match = min(best_matches, key=len)
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
            report_lines.append(f"- يتضح من توزيع البيانات أن هناك توازناً في القيم الرقمية، مما يستدعي مراقبة الانحرافات المعيارية لضمان استقرار الأداء.")
        else:
            top_val = df[col].mode()[0]
            count_top = (df[col] == top_val).sum()
            total = len(df)
            perc = (count_top / total) * 100
            report_lines.append(f"- بعد معالجة البيانات وتوحيد المسميات، تبين أن الفئة الأكثر تكراراً هي '{top_val}'، بنسبة {perc:.1f}% من إجمالي السجلات.")
            report_lines.append(f"- هذا التركز يتطلب دراسة أسباب تفوق هذه الفئة ووضع استراتيجيات تتناسب مع هذا الانتشار.")
        report_lines.append("-" * 30)
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
        if not selected_cols:
            st.error("الرجاء اختيار عمود واحد على الأقل.")
        else:
            with st.spinner('جاري التحليل وتوليد ملف الإكسل المطور...'):
                try:
                    # 1. المعالجة
                    df_processed = df[selected_cols].copy()
                    for col in df_processed.select_dtypes(include=['object']):
                        df_processed[col] = unify_names(df_processed[col])
                    
                    # 2. توليد التقرير
                    full_report = generate_long_report(df_processed, selected_cols)
                    
                    # 3. إنشاء ملف Excel
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
                    for i, col in enumerate(selected_cols):
                        # تنظيف اسم ورقة البيانات المختصرة لتجنب الخطأ
                        summary_sheet_name = clean_sheet_name(col, i)
                        
                        if df_processed[col].dtype in ['int64', 'float64']:
                            summary = df_processed[col].value_counts(bins=5).sort_index().reset_index()
                            chart_type = 'column'
                        else:
                            summary = df_processed[col].value_counts().head(10).reset_index()
                            chart_type = 'pie' if len(summary) <= 5 else 'bar'
                        
                        summary.columns = ['الفئة', 'التكرار']
                        summary.to_excel(writer, sheet_name=summary_sheet_name, index=False)
                        
                        # إنشاء الرسم البياني
                        chart = workbook.add_chart({'type': chart_type})
                        chart.add_series({
                            'name':       f'تحليل {col}',
                            'categories': f"='{summary_sheet_name}'!$A$2:$A${len(summary)+1}",
                            'values':     f"='{summary_sheet_name}'!$B$2:$B${len(summary)+1}",
                        })
                        chart.set_title({'name': f'توزيع: {col}'})
                        
                        # إدراج الرسم في صفحة الداشبورد
                        dashboard_sheet.insert_chart(f'B{row_cursor}', chart)
                        row_cursor += 20 

                    writer.close()
                    
                    st.success("تم التغلب على مشكلة الأسماء وتجهيز الملف بنجاح!")
                    st.download_button(
                        label="📥 تحميل ملف الإكسل (الرسوم قابلة للتعديل)",
                        data=output.getvalue(),
                        file_name="Executive_Analysis_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"حدث خطأ غير متوقع: {e}")
