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
# عنوان الموقع
st.title("📊 محلل البيانات العربي الذكي")
st.write("قم برفع ملف الإكسل، وسنقوم بالباقي!")

uploaded_file = st.file_uploader("اختر ملف إكسل", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("اختر ورقة العمل:", xl.sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    
    st.write("### استعراض البيانات")
    st.dataframe(df.head())

    # اختيار الأعمدة
    all_columns = df.columns.tolist()
    selected_cols = st.multiselect("اختر الأعمدة للتحليل (أو اتركها فارغة لتحليل الكل):", all_columns)
    
    if not selected_cols:
        selected_cols = all_columns

    if st.button("بدء المعالجة والتحليل"):
        with st.spinner('جاري تنظيف وتوحيد البيانات...'):
            df_processed = df[selected_cols].copy()
            
            # عملية التوحيد التلقائي للأعمدة النصية
            for col in df_processed.select_dtypes(include=['object']):
                df_processed[col] = unify_names(df_processed[col])

            st.success("تم توحيد البيانات وتنظيفها!")
            
            # --- قسم الداشبورد ---
            st.header("📈 لوحة المعلومات (Dashboard)")
            cols = st.columns(2)
            charts_images = [] # لتخزين الصور لملف الوورد

            for i, col_name in enumerate(selected_cols):
                target_col = cols[i % 2]
                if df_processed[col_name].dtype in ['int64', 'float64']:
                    # رسم بياني عددي
                    fig = px.histogram(df_processed, x=col_name, title=f"توزيع {col_name}", color_discrete_sequence=['#636EFA'])
                    target_col.plotly_chart(fig)
                else:
                    # رسم بياني فئوي
                    top_values = df_processed[col_name].value_counts().head(10)
                    fig = px.bar(x=top_values.index, y=top_values.values, title=f"أعلى القيم في {col_name}")
                    target_col.plotly_chart(fig)
                
                # حفظ الصورة للتقرير (بشكل مؤقت)
                img_bytes = fig.to_image(format="png")
                charts_images.append((col_name, img_bytes))

            # --- إنشاء التعليقات والتحليل (Textual Analysis) ---
            st.header("📝 التفسير والتحليل")
            analysis_text = []
            for col_name in selected_cols:
                if df_processed[col_name].dtype in ['int64', 'float64']:
                    mean_val = df_processed[col_name].mean()
                    comment = f"العمود '{col_name}': يبلغ متوسط القيم حوالي {mean_val:.2f}. "
                    if df_processed[col_name].skew() > 0:
                        comment += "البيانات تنحو نحو القيم المرتفعة."
                    else:
                        comment += "البيانات موزعة بشكل متوازن نسبيًا."
                else:
                    top_val = df_processed[col_name].mode()[0]
                    comment = f"العمود '{col_name}': القيمة الأكثر تكراراً هي '{top_val}'. "
                
                st.write(f"- {comment}")
                analysis_text.append(comment)

            # --- إنشاء ملفات المخرجات ---
            # 1. ملف إكسل المعالج
            excel_buffer = io.BytesIO()
            df_processed.to_excel(excel_buffer, index=False)
            
            # 2. ملف الوورد
            doc = Document()
            doc.add_heading('تقرير تحليل البيانات التلقائي', 0)
            for text, (col_name, img) in zip(analysis_text, charts_images):
                doc.add_heading(f'تحليل عمود: {col_name}', level=1)
                doc.add_paragraph(text)
                image_stream = io.BytesIO(img)
                doc.add_picture(image_stream, width=Inches(5))
            
            word_buffer = io.BytesIO()
            doc.save(word_buffer)

            # 3. إنشاء ملف ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr("Cleaned_Data.xlsx", excel_buffer.getvalue())
                zip_file.writestr("Analysis_Report.docx", word_buffer.getvalue())

            st.download_button(
                label="📥 تحميل الملفات (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Data_Analysis_Package.zip",
                mime="application/zip"
            )
