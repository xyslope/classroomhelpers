from docx import Document

try:
    doc = Document("toc_example.docx")
    print("ファイルは正常に読み込めました。")
except Exception as e:
    print(f"エラー: {e}")
