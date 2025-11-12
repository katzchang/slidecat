# slidecat

PowerPointファイルをページごとに分割したり結合したりするCLIツール

## 機能

- **分割**: 1つのPPTXファイルを個別のスライドに分割
- **結合**: 複数のPPTXファイルを1つに結合
- **抽出**: 指定した範囲のスライドを抽出

## インストール

```bash
# 依存関係のインストール
pip install -r requirements.txt

# または開発モードでインストール
pip install -e .
```

## 使い方

### 分割 (split)

PPTXファイルをスライドごとに分割します。

```bash
# 1スライドごとに分割（デフォルト）
slidecat split presentation.pptx

# 出力先ディレクトリを指定
slidecat split presentation.pptx --output-dir ./output
slidecat split presentation.pptx -o ./output

# 指定した数のスライドごとに分割
slidecat split presentation.pptx --chunk-size 3
slidecat split presentation.pptx -c 5 -o ./output
```

出力例:
```
# 1スライドごと（デフォルト）
slides/
├── presentation_slide_001.pptx
├── presentation_slide_002.pptx
├── presentation_slide_003.pptx
└── ...

# 3スライドごと (--chunk-size 3)
slides/
├── presentation_slides_001-003.pptx
├── presentation_slides_004-006.pptx
├── presentation_slides_007-009.pptx
└── ...
```

### 結合 (merge)

複数のPPTXファイルを1つに結合します。

```bash
slidecat merge file1.pptx file2.pptx file3.pptx --output combined.pptx
slidecat merge *.pptx -o combined.pptx
```

### 抽出 (extract)

指定した範囲のスライドを抽出します。

```bash
# スライド1〜5を抽出
slidecat extract presentation.pptx --range 1-5 --output extract.pptx

# スライド3から最後まで抽出
slidecat extract presentation.pptx --range 3- --output extract.pptx

# 短縮形
slidecat extract presentation.pptx -r 1-5 -o extract.pptx
```

## コマンドオプション

### `slidecat split`

```
Usage: slidecat split [OPTIONS] INPUT_FILE

Options:
  -o, --output-dir PATH  出力ディレクトリ (デフォルト: ./slides)
  -c, --chunk-size INT   1ファイルあたりのスライド数 (デフォルト: 1)
  --help                 ヘルプを表示
```

### `slidecat merge`

```
Usage: slidecat merge [OPTIONS] INPUT_FILES...

Options:
  -o, --output PATH  出力ファイルパス (必須)
  --help             ヘルプを表示
```

### `slidecat extract`

```
Usage: slidecat extract [OPTIONS] INPUT_FILE

Options:
  -o, --output PATH  出力ファイルパス (必須)
  -r, --range TEXT   抽出するスライド範囲 (例: "1-5" or "3-") (必須)
  --help             ヘルプを表示
```

## 開発

### プロジェクト構造

```
slidecat/
├── slidecat/
│   ├── __init__.py    # パッケージ情報
│   ├── cli.py         # CLIインターフェース
│   └── core.py        # コア機能
├── pyproject.toml     # プロジェクト設定
├── requirements.txt   # 依存関係
└── README.md          # このファイル
```

### 依存関係

- **python-pptx**: PowerPointファイルの読み書き
- **click**: CLIフレームワーク

## ライセンス

MIT

## 注意事項

- 入力ファイルは `.pptx` 形式である必要があります（古い `.ppt` 形式は非対応）
- スライドの分割時、複雑なアニメーションやマクロは保持されない場合があります
- 大きなファイルの処理には時間がかかる場合があります
