import os
import re
from docx import Document

def file_name(text):
    # Replace " - " and "/" with "-"
    text = re.sub(r"\s*–\s*|\s*-\s*|\s*/\s*", "-", text)
    
    # Replace spaces with "_"
    text = text.replace(" ", "_")
    
    # Remove \:*?"<>|
    text = re.sub(r'[\\:*?"<>|]', '', text)
    
    # Remove accents from Vietnamese letters
    mapping = {
        'á': 'a', 'à': 'a', 'ả': 'a', 'ã': 'a', 'ạ': 'a', 'ă': 'a', 'ắ': 'a', 'ằ': 'a', 'ẳ': 'a', 'ẵ': 'a', 'ặ': 'a', 'â': 'a', 'ấ': 'a', 'ầ': 'a', 'ẩ': 'a', 'ẫ': 'a', 'ậ': 'a',
        'đ': 'd',
        'é': 'e', 'è': 'e', 'ẻ': 'e', 'ẽ': 'e', 'ẹ': 'e', 'ê': 'e', 'ế': 'e', 'ề': 'e', 'ể': 'e', 'ễ': 'e', 'ệ': 'e',
        'í': 'i', 'ì': 'i', 'ỉ': 'i', 'ĩ': 'i', 'ị': 'i',
        'ó': 'o', 'ò': 'o', 'ỏ': 'o', 'õ': 'o', 'ọ': 'o', 'ô': 'o', 'ố': 'o', 'ồ': 'o', 'ổ': 'o', 'ỗ': 'o', 'ộ': 'o', 'ơ': 'o', 'ớ': 'o', 'ờ': 'o', 'ở': 'o', 'ỡ': 'o', 'ợ': 'o',
        'ú': 'u', 'ù': 'u', 'ủ': 'u', 'ũ': 'u', 'ụ': 'u', 'ư': 'u', 'ứ': 'u', 'ừ': 'u', 'ử': 'u', 'ữ': 'u', 'ự': 'u',
        'ý': 'y', 'ỳ': 'y', 'ỷ': 'y', 'ỹ': 'y', 'ỵ': 'y'
    }
    text = ''.join(mapping.get(char, char) for char in text)
    
    return text

def doc2txt(data_dir):
    doc = Document(data_dir)

    # Initialize variables
    current_report = None
    current_paragraph = None

    # Iterate over paragraphs
    for paragraph in doc.paragraphs:
        # Retrieve the text without formatting
        text = paragraph.text.strip()

        # Reset if empty line
        if not text:
            current_report = None
            current_paragraph = None
            file.close()
        # Check if the paragraph is a new report's title
        elif not current_report:
            current_report = text
            txt_dir = f'{cwd}\\{file_name(text)}.txt'
            # Create a new text file for the current report
            with open(txt_dir, 'w', encoding='utf-8') as file:
                file.write(f'==={current_report.upper()}===\n')
        # Check if the paragraph is a paragraph's subtitle
        elif paragraph.runs[0].bold:
            current_paragraph = text
            with open(txt_dir, 'a', encoding='utf-8') as file:
                file.write(f'\n=={current_paragraph.upper()}==\n')
        # Check if the paragraph is a new paragraph's content
        else:
            if not current_paragraph:
                current_paragraph = current_report
                with open(txt_dir, 'a', encoding='utf-8') as file:
                    file.write(f'\n=={current_paragraph.upper()}==\n')
            with open(txt_dir, 'a', encoding='utf-8') as file:
                file.write(f'{text}\n')

if __name__ == '__main__':
    cwd = os.getcwd()
    data_dir = 'D:\\Downloads\\Template_phan_tich_doanh_nghiep.docx'
    doc2txt(data_dir)
    
