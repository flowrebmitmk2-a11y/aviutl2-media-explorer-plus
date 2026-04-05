# MediaExplorerPlus

`MediaExplorerPlus` は AviUtl2 用のプラグインです。素材フォルダをタブ付きで開き、メディア探索をしやすくします。

## 導入方法

1. `MediaExplorerPlus.aux2` を AviUtl2 のプラグイン配置先へ置きます。
2. 必要に応じて AviUtl2 を再起動します。

## AviUtl2 上での簡単な使い方

1. AviUtl2 で `Media Explorer Plus` のウィンドウを開きます。
2. 上部の戻る、進む、上へボタンやアドレス欄を使ってフォルダを移動します。
3. 大タブ、小タブを切り替えてよく使う場所をまとめます。
4. 右側のボタンからコピー、貼り付け、alias 保存などを使います。

## ディレクトリ構成

- `src/`: プラグイン本体
- `external/aviutl2-plugin-sdk/`: AviUtl2 SDK の git submodule
- `scripts/build.ps1`: ローカルビルド用スクリプト
- `.github/workflows/build.yml`: CI ビルド

## SDK の入れ方

この repo は `aviutl2/aviutl2_sdk_mirror` を submodule として入れる前提です。

```powershell
git submodule add https://github.com/aviutl2/aviutl2_sdk_mirror.git external/aviutl2-plugin-sdk
git submodule update --init --recursive
```

## ビルド

```powershell
pwsh ./scripts/build.ps1
```

CMake を直接使う場合:

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Release
```

生成物は `build/Release/MediaExplorerPlus.aux2` です。

## GitHub Actions

次のタイミングでビルドが走ります。

- push
- pull request
- 手動実行
- `repository_dispatch` の `sdk-updated`
