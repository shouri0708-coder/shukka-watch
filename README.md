# 出荷調整リアルタイム状況

厚生労働省「医療用医薬品供給状況」Excel を定期取得し、限定出荷・供給停止の
新規・継続・解除を差分表示するページ。GitHub Actions が平日 9/12/15/18 時(JST)に更新。

- 表示: `docs/index.html`（GitHub Pages）
- 自店採用品目: `watch_list.txt` に YJコード or 品名の一部を1行ずつ
- 手動更新: Actions → update → Run workflow
