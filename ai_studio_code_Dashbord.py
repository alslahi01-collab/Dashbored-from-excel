import streamlit as st
import pandas as pd
import io
from rapidfuzz import process
import xlsxwriter

st.set_page_config(page_title="المحلل التنفيذي المتقدم", layout="wide")

# --- دالة توحيد المسميات الذكية ---
def unify_names(series):
    series = series.fillna("غير محدد").astype(str).str.strip()
    if series.nunique() <= 1:
        return series
    
    counts = series.value_counts()
    unique_names = counts.index.tolist()
    mapping = {}
    processed = set()
    
    for name in unique_names:
        if name in processed: continue
        # البحث عن الأسماء المشابهة بنسبة 85%
        matches = process.extract(name, unique_names, score_cutoff=85)
        group = [m[0] for m in matches]
        
        # اختيار الأكثر تكراراً، وعند التساوي اختيار الأقصر
        max_freq = counts[group].max()
        best_matches = [n for n in group if counts[n] == max_freq]
        best_match = min(best_matches, key=len)
        
        for m in group:
            mapping[m] = best_match
            processed.add(m)
    return series.map(mapping)

# --- دالة إنشاء التقرير التنفيذي المفصل ---
def generate_long_report(df, selected_cols):
    report_lines = [
        "تقرير تحليل البيانات الشامل وإرشادات أصحاب المصلحة",
        "================================================",
        f"تاريخ التقرير: {pd.Timestamp.now().strftime('%Y-%m-%d')}",
        f"عدد السجلات الإجمالية: {len(df)} سجل",
        ""
    ]
    
    for col in selected_cols:
        if df[col].dropna().empty:
            continue
            
        report_lines.append(f"🔍 المحور التحليلي: {col}")
        
        if pd.api.types.is_numeric_dtype(df[col]):
            desc = df[col].describe()
            report_lines.append(f"   • نظرة إحصائية: يظهر هذا العمود متوسطاً حسابياً قدره {desc['mean']:.2f}.")
            report_lines.append(f"   • النطاق: تتراوح القيم بين حد أدنى {desc['min']} وحد أقصى {desc['max']}.")
            report_lines.append(f"   • التوصية التنفيذية: تظهر البيانات تشتتاً بمقدار {desc['std']:.2f}. يجب على فريق العمل مراقبة الانحرافات التي تتجاوز القيم المعيارية.")
        else:
            top_values = df[col].value_counts().head(3)
            top_val = top_values.index[0]
            perc = (top_values.values[0] / len(df)) * 100
            report_lines.append(f"   • التحليل النوعي: بعد توحيد المسميات، تبين أن '{top_val}' هي القيمة الأكثر شيوعاً بنسبة {perc:.1f}%.")
            report_lines.append(f"   • التوزيع: تم رصد {df[col].nunique()} فئات مختلفة في هذا المحور.")
            report_lines.append(f"   • رؤية تحليلية: نوصي بتركيز الموارد على الفئات الأكثر تكراراً لرفع كفاءة الاستجابة.")
        
        report_lines.append("-" * 30)
    
    return "\n".join(report_lines)

# --- الواجهة الرسومية ---
st.title("🚀 نظام تحليل البيانات وإعداد التقارير التنفيذية")
st.markdown("قم برفع ملفك وسيقوم النظام بمعالجة البيانات، توحيد الأسماء، وإنشاء داشبورد إكسل تفاعلي.")

uploaded_file = st.file_uploader("ارفع ملف الإكسل", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("اختر ورقة العمل:", xl.sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    
    # فلترة الأعمدة التي لا تحتوي على بيانات
    valid_columns = [col for col in df.columns if not df[col].dropna().empty]
    selected_cols = st.multiselect("اختر الأعمدة للتحليل:", valid_columns, default=valid_columns[:10]) # الافتراضي أول 10 لتجنب البطء

    if st.button("توليد التقرير والرسوم البيانية"):
        if not selected_cols:
            st.error("الرجاء اختيار عمود واحد على الأقل.")
        else:
            with st.spinner('جاري التحليل العميق وتوليد ملف الإكسل المتقدم...'):
                df_processed = df[selected_cols].copy()
                
                # معالجة النصوص فقط
                for col in df_processed.select_dtypes(include=['object']):
                    df_processed[col] = unify_names(df_processed[col])
                
                full_report = generate_long_report(df_processed, selected_cols)
                
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                
                # 1. صفحة البيانات
                df_processed.to_excel(writer, sheet_name='البيانات المعالجة', index=False)
                
                # 2. صفحة التقرير
                report_df = pd.DataFrame({'التقرير التنفيذي': full_report.split('\n')})
                report_df.to_excel(writer, sheet_name='التقرير التحليلي', index=False)
                
                # 3. صفحة الداشبورد
                workbook = writer.book
                dashboard_sheet = workbook.add_worksheet('لوحة المعلومات')
                row_cursor = 1
                
                for col in selected_cols:
                    try:
                        # إنشاء ورقة بيانات للرسم (محدودة بـ 31 حرف لاسم الورقة)
                        clean_col_name = str(col)[:25].replace('[','').replace(']','')
                        summary_sheet_name = f"Data_{clean_col_name}"
                        
                        if pd.api.types.is_numeric_dtype(df_processed[col]):
                            # إذا كانت البيانات رقمية ولها قيم كثيرة نستخدم التقسيم (مع معالجة الخطأ)
                            if df_processed[col].nunique() > 10:
                                summary = df_processed[col].value_counts(bins=5).sort_index().reset_index()
                                summary.columns = ['الفئة', 'التكرار']
                                summary['الفئة'] = summary['الفئة'].astype(str)
                                chart_type = 'column'
                            else:
                                summary = df_processed[col].value_counts().sort_index().reset_index()
                                summary.columns = ['الفئة', 'التكرار']
                                chart_type = 'column'
                        else:
                            summary = df_processed[col].value_counts().head(10).reset_index()
                            summary.columns = ['الفئة', 'التكرار']
                            chart_type = 'pie' if len(summary) <= 4 else 'bar'

                        if summary.empty: continue

                        summary.to_excel(writer, sheet_name=summary_sheet_name, index=False)
                        
                        chart = workbook.add_chart({'type': chart_type})
                        chart.add_series({
                            'name':       f'تحليل {col}',
                            'categories': f"='{summary_sheet_name}'!$A$2:$A${len(summary)+1}",
                            'values':     f"='{summary_sheet_name}'!$B$2:$B${len(summary)+1}",
                        })
                        chart.set_title({'name': f'توزيع: {col}'})
                        dashboard_sheet.insert_chart(f'B{row_cursor}', chart)
                        row_cursor += 20
                    except:
                        continue # في حال فشل رسم عمود معين، ننتقل للذي يليه

                writer.close()
                
                st.success("✅ تم بنجاح معالجة البيانات!")
                st.download_button(
                    label="📥 تحميل ملف الإكسل (تحليل كامل + رسوم قابلة للتعديل)",
                    data=output.getvalue(),
                    file_name="Advanced_Executive_Analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                with st.expander("عرض التقرير التنفيذي المكتوب"):
                    st.text(full_report)
