import streamlit as st
import pandas as pd
import io
from rapidfuzz import process
import xlsxwriter

# إعداد الصفحة
st.set_page_config(page_title="المحلل الاحترافي", layout="wide")

st.title("📊 محلل البيانات التنفيذي")
st.markdown("قم برفع ملف الإكسل وسيتم تنظيفه وتحليله تلقائياً.")

# 1. تحميل الملف
uploaded_file = st.file_uploader("الخطوة 1: ارفع ملف الإكسل هنا", type=["xlsx"])

if uploaded_file:
    try:
        # قراءة أسماء أوراق العمل
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        # 2. اختيار ورقة العمل
        selected_sheet = st.selectbox("الخطوة 2: اختر ورقة العمل التي تريد تحليلها:", sheet_names)
        
        # تحميل البيانات
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
        if df.empty:
            st.warning("هذه الورقة فارغة، يرجى اختيار ورقة عمل أخرى.")
        else:
            st.success(f"تم تحميل {len(df)} سجل بنجاح.")
            
            # 3. اختيار الأعمدة
            all_cols = df.columns.tolist()
            st.info("الخطوة 3: اختر الأعمدة (تم تحديد أول 10 أعمدة تلقائياً لتسريع العملية)")
            selected_cols = st.multiselect("الأعمدة المختارة:", all_cols, default=all_cols[:10])
            
            if not selected_cols:
                st.warning("يرجى اختيار عمود واحد على الأقل للبدء.")
            else:
                # 4. زر البدء
                if st.button("🚀 ابدأ التحليل العميق وتوليد التقرير"):
                    with st.spinner('جاري تنظيف البيانات وتوليد الرسوم البيانية...'):
                        
                        # --- معالجة البيانات ---
                        df_final = df[selected_cols].copy()
                        
                        # توحيد المسميات (Logic)
                        for col in df_final.select_dtypes(include=['object']):
                            # تنظيف النصوص من الفراغات
                            df_final[col] = df_final[col].astype(str).str.strip()
                            
                            counts = df_final[col].value_counts()
                            unique_vals = counts.index.tolist()
                            mapping = {}
                            processed_set = set()
                            
                            for val in unique_vals:
                                if val in processed_set: continue
                                matches = process.extract(val, unique_vals, score_cutoff=85)
                                group = [m[0] for m in matches]
                                
                                # الأكثر تكراراً
                                max_f = counts[group].max()
                                winners = [n for n in group if counts[n] == max_f]
                                # الأقصر طولاً
                                best = min(winners, key=len)
                                
                                for m in group:
                                    mapping[m] = best
                                    processed_set.add(m)
                            
                            df_final[col] = df_final[col].map(mapping)

                        # --- توليد التقرير النصي ---
                        report_text = [
                            "تقرير تحليل البيانات الاستراتيجي",
                            "==========================",
                            f"إجمالي السجلات المعالجة: {len(df_final)}",
                            ""
                        ]
                        
                        for col in selected_cols:
                            if df_final[col].dtype in ['int64', 'float64']:
                                m = df_final[col].mean()
                                report_text.append(f"• عمود {col}: المتوسط الحسابي هو {m:.2f}. نوصي بمتابعة القيم التي تبتعد عن هذا المتوسط.")
                            else:
                                top = df_final[col].mode()[0]
                                report_text.append(f"• عمود {col}: القيمة الأكثر شيوعاً هي '{top}'. هذا المؤشر يعكس الاتجاه العام في هذا المحور.")

                        # --- إنشاء ملف Excel المتقدم ---
                        output = io.BytesIO()
                        writer = pd.ExcelWriter(output, engine='xlsxwriter')
                        
                        # ورقة البيانات
                        df_final.to_excel(writer, sheet_name='البيانات المعدلة', index=False)
                        
                        # ورقة التقرير
                        report_df = pd.DataFrame(report_text, columns=['التحليل التنفيذي'])
                        report_df.to_excel(writer, sheet_name='التقرير', index=False)
                        
                        # ورقة الداشبورد والرسوم
                        workbook = writer.book
                        dash_sheet = workbook.add_worksheet('الرسومات البيانية')
                        
                        row_idx = 1
                        for col in selected_cols:
                            try:
                                # تجهيز بيانات الرسم
                                data_sheet_name = f"D_{str(col)[:25]}"
                                if df_final[col].dtype in ['int64', 'float64']:
                                    summary = df_final[col].value_counts(bins=min(5, df_final[col].nunique())).sort_index()
                                    c_type = 'column'
                                else:
                                    summary = df_final[col].value_counts().head(10)
                                    c_type = 'pie' if len(summary) < 5 else 'bar'
                                
                                summary.to_excel(writer, sheet_name=data_sheet_name)
                                
                                # رسم بياني أصلي في إكسل
                                chart = workbook.add_chart({'type': c_type})
                                chart.add_series({
                                    'name': col,
                                    'categories': f"='{data_sheet_name}'!$A$2:$A${len(summary)+1}",
                                    'values':     f"='{data_sheet_name}'!$B$2:$B${len(summary)+1}",
                                })
                                chart.set_title({'name': f'تحليل {col}'})
                                dash_sheet.insert_chart(f'B{row_idx}', chart)
                                row_idx += 20
                            except:
                                continue
                        
                        writer.close()
                        
                        # --- عرض النتائج والتحميل ---
                        st.success("🏁 اكتمل التحليل!")
                        st.download_button(
                            label="📥 تحميل التقرير النهائي (Excel المطور)",
                            data=output.getvalue(),
                            file_name="Executive_Dashboard.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.text_area("ملخص التقرير:", "\n".join(report_text), height=200)

    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
