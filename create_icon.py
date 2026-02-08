"""
画像ファイル（PNG, JPG）からアイコンファイル（.ico）を作成するスクリプト

使用方法:
    python create_icon.py

推奨画像サイズ: 256x256ピクセル以上
"""

import sys
from tkinter import Tk, filedialog, messagebox
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    print("エラー: Pillowがインストールされていません")
    print("以下のコマンドを実行してください:")
    print("  pip install Pillow")
    sys.exit(1)

def create_icon_from_image(image_path, output_path=None):
    """画像ファイルからアイコンを作成"""
    try:
        # 画像を開く
        img = Image.open(image_path)
        
        # RGBA モードに変換（透過対応）
        if img.mode != 'RGBA':
            img = img.convert('RGBA')
        
        # 出力パスが指定されていない場合は自動生成
        if output_path is None:
            output_path = Path(image_path).with_suffix('.ico')
        
        # 複数サイズのアイコンを生成（16x16, 32x32, 48x48, 64x64, 128x128, 256x256）
        icon_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
        
        img.save(
            output_path,
            format='ICO',
            sizes=icon_sizes
        )
        
        return True, output_path
    
    except Exception as e:
        return False, str(e)

def main():
    print("="*60)
    print("画像からアイコンファイル(.ico)を作成")
    print("="*60)
    
    # Tkinter初期化
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # 画像ファイルを選択
    messagebox.showinfo(
        'アイコン作成',
        'アイコンにしたい画像ファイルを選択してください。\n\n'
        '推奨サイズ: 256x256ピクセル以上\n'
        '対応形式: PNG, JPG, BMP, GIF など'
    )
    
    image_path = filedialog.askopenfilename(
        title='画像ファイルを選択',
        filetypes=[
            ('画像ファイル', '*.png *.jpg *.jpeg *.bmp *.gif *.tiff'),
            ('すべてのファイル', '*.*')
        ]
    )
    
    if not image_path:
        messagebox.showwarning('キャンセル', '画像が選択されませんでした。')
        root.destroy()
        return
    
    print(f"\n選択された画像: {image_path}")
    
    # 保存先を選択
    default_name = Path(image_path).stem + '.ico'
    output_path = filedialog.asksaveasfilename(
        title='アイコンファイルの保存先',
        defaultextension='.ico',
        initialfile=default_name,
        filetypes=[('アイコンファイル', '*.ico')]
    )
    
    if not output_path:
        messagebox.showwarning('キャンセル', '保存先が選択されませんでした。')
        root.destroy()
        return
    
    root.destroy()
    
    # アイコン作成
    print(f"アイコンを作成中...")
    success, result = create_icon_from_image(image_path, output_path)
    
    if success:
        print(f"\n✓ アイコンファイルを作成しました:")
        print(f"  {result}")
        print(f"\nこのアイコンをビルド時に使用できます:")
        print(f"  1. 'icon.ico' という名前で保存すると自動認識されます")
        print(f"  2. または python build_exe.py 実行時に選択します")
        
        # 成功メッセージ
        root2 = Tk()
        root2.withdraw()
        root2.attributes('-topmost', True)
        messagebox.showinfo(
            '完了',
            f'アイコンファイルを作成しました:\n\n{result}\n\n'
            'このアイコンをビルド時に使用できます。'
        )
        root2.destroy()
    else:
        print(f"\n✗ エラー: {result}")
        root2 = Tk()
        root2.withdraw()
        root2.attributes('-topmost', True)
        messagebox.showerror('エラー', f'アイコン作成に失敗しました:\n\n{result}')
        root2.destroy()

if __name__ == '__main__':
    main()
