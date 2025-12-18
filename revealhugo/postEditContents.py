import os
import re

def remove_notes_and_qa_from_markdown(path):
    """MarkdownファイルからスピーカーノートとQ&Aスライドを削除"""
    modified_count = 0
    
    for root, dirs, files in os.walk(path):
        for file in files:
            if file == "_index.md":
                filepath = os.path.join(root, file)
                try:
                    with open(filepath, "r", encoding="utf-8") as f:
                        content = f.read()
                    
                    original_content = content
                    
                    # 1. {{< note >}}...{{< /note >}} を削除
                    content = re.sub(
                        r'\{\{<\s*note\s*>\}\}.*?\{\{<\s*/note\s*>\}\}',
                        '',
                        content,
                        flags=re.DOTALL | re.IGNORECASE
                    )
                    
                    # 2. Q&Aスライドを削除
                    # スライド単位で処理：\n---\n で分割
                    slides = content.split('\n---\n')
                    filtered_slides = []
                    
                    for i, slide in enumerate(slides):
                        # スライドの最初の有効な行をチェック
                        first_line = slide.strip().split('\n')[0] if slide.strip() else ''
                        
                        # # Q&A または ## Q&A で始まるかチェック
                        if re.match(r'^#{1,3}\s+Q&A\s*$', first_line):
                            print(f"[REMOVED] 1 Sheet")
                            # Q&Aスライドの場合
                            # {{% /section %}} が含まれていたら、それだけ残す
                            if '{{% /section %}}' in slide:
                                filtered_slides.append('{{% /section %}}')
                            # そうでなければこのスライドは削除（appendしない）
                        else:
                            # Q&Aでない場合は保持
                            filtered_slides.append(slide)
                    
                    # スライドを再結合
                    content = '\n---\n'.join(filtered_slides)
                    
                    # 3. 連続する空行を整理
                    content = re.sub(r'\n{3,}', '\n\n', content)
                    
                    # ファイルの末尾の空行を統一
                    content = content.rstrip() + '\n'
                    
                    # 変更があった場合のみ書き込み
                    if content != original_content:
                        with open(filepath, "w", encoding="utf-8") as f:
                            f.write(content)
                        print(f"[MODIFIED] {filepath}")
                        modified_count += 1
                    else:
                        print(f"[OK] {filepath}")
                        
                except Exception as e:
                    print(f"[ERROR] {filepath}: {e}")
                    return 1
    
    print(f"\nTotal modified files: {modified_count}")
    return 0

if __name__ == "__main__":
    import sys
    print(f"Current working directory: {os.getcwd()}")
    print(f"Looking for path: content")
    print(f"Full path would be: {os.path.abspath('content')}")
    sys.exit(remove_notes_and_qa_from_markdown("content"))
