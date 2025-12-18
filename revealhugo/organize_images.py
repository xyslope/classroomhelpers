#!/usr/bin/env python3
"""
Reveal-Hugo講義資料の画像整理スクリプト（WMF変換改良版）

使用方法:
    python organize_images_v4.py [--dry]

新機能:
- HTMLのimgタグにも対応
- WMFファイルを複数の方法で自動PNG変換
- より多くの画像形式をサポート

機能:
1. カレントディレクトリの_index.md以外の.mdファイルを抽出
2. 各ファイルを処理して画像を整理
3. ファイル名（拡張子なし）と同名のフォルダを作成
4. 画像リンクを抽出して新フォルダに移動（ファイル名は body_nn.拡張子、nnは2桁通し番号）
5. 画像リンクを新しいパスに修正
6. 二回目以降の実行では、まだ処理されていない画像のみを処理
7. WMFファイルを複数の方法でPNGに変換

WMF変換方法:
1. PIL + Windows API (推奨)
2. wand (ImageMagick)
3. LibreOffice経由変換
4. 手動コピー（変換失敗時）

対応形式:
- Markdown: ![alt](path)
- HTML: <img src="path" ...>
- 画像形式: jpg, jpeg, png, gif, svg, webp, wmf

オプション:
    --dry : 実際の処理は行わず、処理内容のみを表示
"""

import os
import re
import shutil
import argparse
import subprocess
from pathlib import Path, PurePath
from typing import List, Tuple, Dict, Optional

# WMF変換用のライブラリ（オプショナル）
PIL_SUPPORT = False
WAND_SUPPORT = False

try:
    from PIL import Image
    PIL_SUPPORT = True
    print("✅ PIL (Pillow) が利用可能です")
except ImportError:
    pass

try:
    from wand.image import Image as WandImage
    WAND_SUPPORT = True
    print("✅ Wand (ImageMagick) が利用可能です")
except ImportError:
    pass

if not PIL_SUPPORT and not WAND_SUPPORT:
    print("⚠️  注意: PIL/Wand ライブラリが見つかりません。WMF変換機能は制限されます。")
    print("   pip install Pillow または pip install Wand でインストールしてください。")


def find_md_files(directory: Path) -> List[Path]:
    """_index.md以外の.mdファイルを抽出"""
    md_files = []
    for file_path in directory.glob("*.md"):
        if file_path.name != "_index.md":
            md_files.append(file_path)
    return sorted(md_files)


def extract_image_links(content: str) -> List[Tuple[str, str, str, str]]:
    """
    MarkdownファイルとHTMLから画像リンクを抽出
    戻り値: [(元のパス, 元のファイル名, 拡張子, マッチした全体), ...]
    """
    image_links = []
    
    # Markdown画像記法: ![alt](path) または ![alt](path "title")
    markdown_pattern = r'!\[.*?\]\(([^)]+)\)'
    markdown_matches = re.finditer(markdown_pattern, content)
    
    for match in markdown_matches:
        full_match = match.group(0)
        image_path = match.group(1).split('"')[0].strip()
        
        # Windowsパスの処理: バックスラッシュをスラッシュに統一
        image_path = image_path.replace('\\', '/')
        
        if is_supported_image(image_path):
            path_obj = PurePath(image_path)
            filename = path_obj.stem
            extension = path_obj.suffix.lower()
            image_links.append((image_path, filename, extension, full_match))
    
    # HTML img タグ: <img src="path" ...>
    html_pattern = r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>'
    html_matches = re.finditer(html_pattern, content, re.IGNORECASE)
    
    for match in html_matches:
        full_match = match.group(0)
        image_path = match.group(1)
        
        # Windowsパスの処理: バックスラッシュをスラッシュに統一
        image_path = image_path.replace('\\', '/')
        
        if is_supported_image(image_path):
            path_obj = PurePath(image_path)
            filename = path_obj.stem
            extension = path_obj.suffix.lower()
            image_links.append((image_path, filename, extension, full_match))
    
    return image_links


def is_supported_image(image_path: str) -> bool:
    """サポートされている画像形式かチェック"""
    supported_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.svg', '.webp', '.wmf']
    return any(image_path.lower().endswith(ext) for ext in supported_extensions)


def convert_wmf_with_pil(source_path: Path, target_path: Path) -> bool:
    """PILを使ってWMFファイルをPNGに変換"""
    if not PIL_SUPPORT:
        return False
    
    try:
        with Image.open(source_path) as img:
            # RGBAモードに変換（透明度対応）
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            
            # PNGとして保存
            img.save(target_path, 'PNG')
            return True
    except Exception as e:
        print(f"      📝 PIL変換失敗: {e}")
        return False


def convert_wmf_with_wand(source_path: Path, target_path: Path) -> bool:
    """Wandを使ってWMFファイルをPNGに変換"""
    if not WAND_SUPPORT:
        return False
    
    try:
        with WandImage(filename=str(source_path)) as img:
            img.format = 'png'
            img.save(filename=str(target_path))
            return True
    except Exception as e:
        print(f"      📝 Wand変換失敗: {e}")
        return False


def convert_wmf_with_libreoffice(source_path: Path, target_path: Path) -> bool:
    """LibreOfficeを使ってWMFファイルをPNGに変換"""
    try:
        # LibreOfficeが利用可能かチェック
        result = subprocess.run(['soffice', '--version'], 
                              capture_output=True, text=True, timeout=5)
        
        if result.returncode != 0:
            return False
        
        # LibreOfficeでWMFをPNGに変換
        temp_dir = target_path.parent
        cmd = [
            'soffice', '--headless', '--convert-to', 'png',
            '--outdir', str(temp_dir), str(source_path)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            # 生成されたファイルを目的の名前にリネーム
            generated_file = temp_dir / f"{source_path.stem}.png"
            if generated_file.exists():
                shutil.move(str(generated_file), str(target_path))
                return True
        
        return False
    except Exception as e:
        print(f"      📝 LibreOffice変換失敗: {e}")
        return False


def convert_wmf_to_png(source_path: Path, target_path: Path) -> bool:
    """WMFファイルをPNGに変換（複数の方法を試行）"""
    print(f"      🔄 WMF変換を試行中...")
    
    # 方法1: PIL
    if convert_wmf_with_pil(source_path, target_path):
        print(f"      ✅ PIL変換成功")
        return True
    
    # 方法2: Wand (ImageMagick)
    if convert_wmf_with_wand(source_path, target_path):
        print(f"      ✅ Wand変換成功")
        return True
    
    # 方法3: LibreOffice
    if convert_wmf_with_libreoffice(source_path, target_path):
        print(f"      ✅ LibreOffice変換成功")
        return True
    
    print(f"      ❌ 全ての変換方法が失敗しました")
    return False


def update_image_links(content: str, old_new_map: Dict[str, Tuple[str, str]]) -> str:
    """画像リンクを新しいパスに置換"""
    updated_content = content
    
    for old_path, (new_path, original_match) in old_new_map.items():
        # Markdownの場合
        if original_match.startswith('!['):
            # Markdown記法の置換
            pattern = re.escape(old_path)
            updated_content = re.sub(pattern, new_path, updated_content)
        
        # HTMLの場合
        elif original_match.startswith('<img'):
            # HTML img タグのsrc属性を置換
            old_src_pattern = re.escape(old_path)
            updated_content = re.sub(old_src_pattern, new_path, updated_content)
        
        # 一般的な置換（Windowsパス対応）
        old_path_normalized = old_path.replace('\\', '/')
        old_path_windows = old_path.replace('/', '\\')
        
        for path_variant in [old_path, old_path_normalized, old_path_windows]:
            pattern = re.escape(path_variant)
            updated_content = re.sub(pattern, new_path, updated_content)
    
    return updated_content


def safe_move_file(source_path: Path, target_path: Path) -> bool:
    """
    ファイルを安全に移動する（Windowsでのパーミッションエラー対応）
    """
    try:
        # 移動先ディレクトリが存在することを確認
        target_path.parent.mkdir(parents=True, exist_ok=True)
        
        # ファイルが既に存在する場合は削除
        if target_path.exists():
            target_path.unlink()
        
        # ファイルを移動
        shutil.move(str(source_path), str(target_path))
        return True
    except PermissionError as e:
        print(f"      ❌ 権限エラー: {e}")
        return False
    except FileNotFoundError as e:
        print(f"      ❌ ファイルが見つかりません: {e}")
        return False
    except Exception as e:
        print(f"      ❌ 移動エラー: {e}")
        return False


def process_wmf_file(source_path: Path, target_path: Path, fallback_copy: bool = True) -> bool:
    """WMFファイルを処理（PNG変換または手動コピー）"""
    try:
        # 移動先ディレクトリが存在することを確認
        target_path.parent.mkdir(parents=True, exist_ok=True)
        
        # ファイルが既に存在する場合は削除
        if target_path.exists():
            target_path.unlink()
        
        # WMFをPNGに変換を試行
        if convert_wmf_to_png(source_path, target_path):
            # 変換成功したら元ファイルを削除
            source_path.unlink()
            return True
        elif fallback_copy:
            # 変換失敗時は拡張子をwmfのままコピー
            wmf_target_path = target_path.with_suffix('.wmf')
            shutil.copy2(str(source_path), str(wmf_target_path))
            print(f"      📋 変換失敗のためWMFファイルをコピー: {wmf_target_path.name}")
            # 元ファイルを削除
            source_path.unlink()
            return True
        else:
            return False
    except Exception as e:
        print(f"      ❌ WMF処理エラー: {e}")
        return False


def get_next_available_index(target_folder: Path, file_body: str) -> int:
    """次に使用可能な通し番号を取得"""
    max_index = 0
    if target_folder.exists():
        for file_path in target_folder.glob(f"{file_body}_*.*"):
            try:
                # ファイル名から番号を抽出: body_nn.ext -> nn
                name_without_ext = file_path.stem
                if name_without_ext.startswith(f"{file_body}_"):
                    index_str = name_without_ext[len(file_body)+1:]
                    if index_str.isdigit():
                        max_index = max(max_index, int(index_str))
            except:
                continue
    return max_index + 1


def is_already_in_target_folder(image_path: str, file_body: str) -> bool:
    """画像が既にターゲットフォルダにあるかチェック"""
    # パスがbody/で始まっているかチェック
    normalized_path = image_path.replace('\\', '/')
    return normalized_path.startswith(f"{file_body}/")


def filter_unprocessed_images(image_links: List[Tuple[str, str, str, str]], file_body: str) -> List[Tuple[str, str, str, str]]:
    """まだ処理されていない画像のみをフィルタリング"""
    unprocessed = []
    for old_path, original_filename, extension, full_match in image_links:
        if not is_already_in_target_folder(old_path, file_body):
            unprocessed.append((old_path, original_filename, extension, full_match))
    return unprocessed


def process_file(md_file: Path, dry_run: bool = False) -> None:
    """個別のmdファイルを処理"""
    print(f"\n📄 処理中: {md_file.name}")
    
    # ファイル名のbody部分を取得（拡張子なし）
    file_body = md_file.stem
    target_folder = md_file.parent / file_body
    
    # フォルダの存在確認
    if target_folder.exists():
        print(f"   📁 フォルダは既に存在: {target_folder}")
        existing_files = list(target_folder.glob(f"{file_body}_*.*"))
        if existing_files:
            print(f"   📋 既存ファイル数: {len(existing_files)}")
    else:
        print(f"   📁 作成予定フォルダ: {target_folder}")
    
    # mdファイルの内容を読み込み
    try:
        with open(md_file, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"   ❌ ファイル読み込みエラー: {e}")
        return
    
    # 画像リンクを抽出（MarkdownとHTML両方）
    all_image_links = extract_image_links(content)
    
    if not all_image_links:
        print(f"   ℹ️  画像リンクが見つかりませんでした")
        return
    
    # まだ処理されていない画像のみをフィルタリング
    unprocessed_images = filter_unprocessed_images(all_image_links, file_body)
    
    print(f"   🖼️  総画像リンク数: {len(all_image_links)}")
    print(f"   🆕 未処理画像数: {len(unprocessed_images)}")
    
    if not unprocessed_images:
        print(f"   ✅ すべての画像は既に処理済みです")
        return
    
    # 次に使用可能な通し番号を取得
    next_index = get_next_available_index(target_folder, file_body)
    print(f"   🔢 開始番号: {next_index:02d}")
    
    old_new_map = {}
    
    for i, (old_path, original_filename, extension, full_match) in enumerate(unprocessed_images):
        # 新しいファイル名を決定
        current_index = next_index + i
        
        # WMFの場合の処理
        if extension.lower() == '.wmf':
            # 変換成功時はPNG、失敗時はWMFのまま
            new_extension_png = '.png'
            new_extension_wmf = '.wmf'
            conversion_note = " (WMF→PNG変換試行)"
        else:
            new_extension_png = extension
            new_extension_wmf = extension
            conversion_note = ""
        
        new_filename_png = f"{file_body}_{current_index:02d}{new_extension_png}"
        new_filename_wmf = f"{file_body}_{current_index:02d}{new_extension_wmf}"
        
        # パスの処理（Windowsとの互換性を考慮）
        if os.path.isabs(old_path):
            source_path = Path(old_path)
        else:
            # スラッシュ区切りのパスをWindowsパスに変換
            normalized_path = old_path.replace('/', os.sep)
            source_path = md_file.parent / normalized_path
        
        target_path_png = target_folder / new_filename_png
        target_path_wmf = target_folder / new_filename_wmf
        
        print(f"      • {old_path}")
        print(f"        └→ {file_body}/{new_filename_png} (#{current_index:02d}){conversion_note}")
        
        success = False
        actual_new_filename = new_filename_png
        
        if not dry_run:
            # フォルダが存在しない場合は作成
            if not target_folder.exists():
                try:
                    target_folder.mkdir(parents=True, exist_ok=True)
                    print(f"      📁 フォルダを作成: {target_folder}")
                except Exception as e:
                    print(f"      ❌ フォルダ作成エラー: {e}")
                    continue
            
            # ファイルを処理
            if source_path.exists():
                if extension.lower() == '.wmf':
                    # WMFファイルの場合は変換を試行
                    if process_wmf_file(source_path, target_path_png, fallback_copy=True):
                        # 実際に作成されたファイルを確認
                        if target_path_png.exists():
                            print(f"      ✅ WMF→PNG変換完了: {new_filename_png}")
                            actual_new_filename = new_filename_png
                        elif target_path_wmf.exists():
                            print(f"      📋 WMFファイルとしてコピー: {new_filename_wmf}")
                            actual_new_filename = new_filename_wmf
                        success = True
                    else:
                        print(f"      ❌ WMF処理失敗: {old_path}")
                        continue
                else:
                    # 通常のファイル移動
                    if safe_move_file(source_path, target_path_png):
                        print(f"      ✅ 移動完了: {new_filename_png}")
                        success = True
                    else:
                        continue
            else:
                print(f"      ⚠️  ソースファイルが見つかりません: {source_path}")
                continue
        else:
            success = True
        
        if success:
            # パス置換マップに追加（スラッシュ区切りで統一）
            new_relative_path = f"{file_body}/{actual_new_filename}"
            old_new_map[old_path] = (new_relative_path, full_match)
    
    # mdファイル内の画像リンクを更新
    if old_new_map:
        updated_content = update_image_links(content, old_new_map)
        
        if not dry_run:
            try:
                # BOMなしUTF-8で保存
                with open(md_file, 'w', encoding='utf-8', newline='\n') as f:
                    f.write(updated_content)
                print(f"   ✅ {md_file.name}のリンクを更新しました（{len(old_new_map)}個のリンク）")
            except Exception as e:
                print(f"   ❌ ファイル更新エラー: {e}")
        else:
            print(f"   📝 {md_file.name}のリンク更新予定（{len(old_new_map)}個のリンク）")
    else:
        print(f"   ℹ️  更新対象のリンクはありませんでした")


def main():
    parser = argparse.ArgumentParser(
        description="Reveal-Hugo講義資料の画像整理スクリプト（WMF変換改良版）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument(
        "--dry", 
        action="store_true", 
        help="実際の処理は行わず、処理内容のみを表示"
    )
    
    args = parser.parse_args()
    
    current_dir = Path.cwd()
    print(f"🚀 画像整理スクリプト（WMF変換改良版）を開始します")
    print(f"📂 作業ディレクトリ: {current_dir}")
    
    if args.dry:
        print("🔍 DRY RUN モード: 実際の処理は行いません")
    
    # _index.md以外の.mdファイルを検索
    md_files = find_md_files(current_dir)
    
    if not md_files:
        print("❌ 処理対象の.mdファイルが見つかりませんでした")
        return
    
    print(f"📋 処理対象ファイル数: {len(md_files)}")
    for md_file in md_files:
        print(f"    • {md_file.name}")
    
    # 各ファイルを処理
    for md_file in md_files:
        process_file(md_file, dry_run=args.dry)
    
    print(f"\n🎉 処理が完了しました！")
    if args.dry:
        print("💡 実際に処理を実行するには --dry オプションを外して再実行してください")


if __name__ == "__main__":
    main()
