# classroomhelpers
授業の管理に便利なツール用の置き場所

一つのツールとして独立させるまでもないのでまとめて公開しておく。

## markdownmatrix：マークダウンを表にする
マークダウン中に表を書くツールではない。

マークダウンでツリー状に書いた文章を表にする。

ルーブリックをマークダウンで書こうと思って作成した。
html, csv, txtに出力可能。
簡単なスクリプトなので、エクスポーターを増やせば、さまざまなスタイルに対応できる。

例えば、こんな文章を用意する。

```
# 見出し１
## 中身１－１
## 中身１－２
## 中身１－３
# 見出し２
## 中身２－１
説明も書ける
## 中身２－２
## 中身２－３
# 見出し３
## 中身３－１
## 中身３－２
## 中身３－３
## 中身３－４
# 見出し４
## 中身４－１
## 中身４－２
## 中身４－３
```

頭の整理時には、こんな感じで思いつくままにツリー状に書くと便利なんだけど、じゃあ、これで全部網羅できてる？とか思うと、少し一覧性が悪い。

そこで、このツールを使うと、次のようになる。

![処理結果の例](sample.png)

htmlで出力したものは[こちら](markdownsource.html)。

現状では、２レベルまでのタイトルを認識して、１レベル目を行（縦）、２レベル目を列（横）に展開した表にする。
１レベル目の数だけ行ができて、１レベルの中にある２レベルの数だけ列ができる。

htmlのばあい、最初の塊（１レベルとその下の２レベル）をヘッダーとする。

ルーブリックに使用する場合、１レベル目が評価軸、２レベル目が評価基準となる。

## createreport：wordファイルをまとめてpdfに変換して一つにまとめる
Adobeなどを使ってもよいが、目次などを作りたい場合に便利。

filelist.csvに作成対象のファイルリストを入れておく。いくつかの列があるが、現状で使用しているのはファイル名のみ。

冊子を作る場合など、奇数ページのファイルに空白ページを加えて、全部を同じ側から始めたいという需要がある。
本ツールでは、オプションで空白ページを自動で加える。

また、各ファイルの先頭ページのページ番号を１ページ目からの通算でリストアップするので、目次ファイルを作るときに便利。

ただし、現状、通しのページ番号を入れる機能はない。

docsフォルダのファイルを自動でリストアップしてpdfを作成することもできるが、ファイルの順序が保証されないので、filelist.csvを使ったほうが簡便な気がする。
今後、目次の自動作成機能などを追加したいので、その際にも利用できるようにしておく。

今後の予定
- [ ] 通しのページ番号を入れる
- [ ] 目次ファイルの自動生成（これはけっこう難しそうな気がする）

## （リンクのみ）語群からキーワードを選択する問題を作成する
語群選択問題を作成するスクリプト。これもgistレベルだけど、けっこう便利。

https://blog.ecofirm.com/post/createquestionsfromwordlist/

※　GoogleClassroomでテストをするようになってしまって、このような問題を使用する機会が減った。
今後どうするか、要検討。

## （公開予定）basicmarket：簡単な市場メカニズムの図を作成
pythonのMatplotLibで簡単な市場メカニズムの図を作成する。

講義用の図を作るベースとして。

gistレベルだけど、よく使うのでここに公開しておく。

この図は、xy軸、需要供給曲線、交点座標の自動抽出、線で囲まれたエリアを塗りつぶす、という機能を使用している。
この図を適宜編集すれば、講義で図を使用する際にかなり楽ができる。



