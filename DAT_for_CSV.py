import pandas as pd


def convert_dat_to_csv(input_file, output_file):
    """
    تحويل ملف DAT إلى CSV

    المعاملات:
    input_file (str): مسار ملف DAT المدخل
    output_file (str): مسار ملف CSV المخرج
    """
    try:
        # قراءة ملف DAT
        # نفترض أن البيانات مفصولة بفواصل، يمكنك تغيير sep حسب تنسيق ملفك
        data = pd.read_csv(input_file, sep=',')

        # حفظ البيانات كملف CSV
        data.to_csv(output_file, index=False)
        print(f"تم تحويل الملف بنجاح إلى: {output_file}")

    except Exception as e:
        print(f"حدث خطأ أثناء التحويل: {str(e)}")

# مثال على الاستخدام
file_name = "invoice/10004.dat"

convert_dat_to_csv(file_name, 'output.csv')