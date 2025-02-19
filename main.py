from pdf2docx import Converter

def pdf_to_word(pdf_file):
    # Đổi tên file DOCX từ tên file PDF
    docx_filename = pdf_file.name.replace('.pdf', '.docx')
    
    # Chuyển đổi PDF sang DOCX
    cv = Converter(pdf_file.name)
    cv.convert(docx_filename, multi_processing=True, start=0, end=None)  # multi_processing=True để tăng tốc quá trình
    cv.close()
    
    return docx_filename

# Ví dụ sử dụng
pdf_file_path = 'test.pdf'  # Thay thế bằng đường dẫn file PDF của bạn
output_docx = pdf_to_word(open(pdf_file_path, 'rb'))

print(f"File DOCX đã được tạo: {output_docx}")
