# MediaExplorerPlus

AviUtl2 用の `.aux2` プラグインです。`media_explorer` の右ペイン機能だけを独立させ、エクスプローラーとタブ管理に絞ったサブウインドウとして使います。

## できること

- エクスプローラー表示
- 大タブ/小タブ管理
- アドレスバー移動
- コピー/切り取り/貼り付け
- 選択オブジェクトの alias 保存

`media_explorer` にあった左ペインのプレビュー埋め込みはありません。

## ビルド

```powershell
cmake -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Release
```

実環境へそのまま反映する場合は `build_and_copy.bat` を使えます。
