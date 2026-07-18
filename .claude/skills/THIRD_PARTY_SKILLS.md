# サードパーティ製スキルの出所一覧

このディレクトリには、外部リポジトリから導入したスキルが含まれる。
更新する場合は元リポジトリから再取得して上書きすること(手動編集しない)。

導入日: 2026-07-18

| スキル | 元リポジトリ | 取得時コミット | ライセンス |
|---|---|---|---|
| brainstorming / dispatching-parallel-agents / executing-plans / finishing-a-development-branch / receiving-code-review / requesting-code-review / subagent-driven-development / systematic-debugging / test-driven-development / using-git-worktrees / using-superpowers / verification-before-completion / writing-plans / writing-skills | [obra/superpowers](https://github.com/obra/superpowers) | d884ae0 | MIT |
| planning-with-files | [OthmanAdi/planning-with-files](https://github.com/OthmanAdi/planning-with-files) | 7c6c6cb | MIT |
| frontend-design | [anthropics/skills](https://github.com/anthropics/skills) | fa0fa64 | Apache-2.0 (LICENSE.txt同梱) |
| understand | [Lum1104/Understand-Anything](https://github.com/Lum1104/Understand-Anything) | b9ac6be | MIT |
| ask-questions-if-underspecified / second-opinion / sharp-edges | [trailofbits/skills](https://github.com/trailofbits/skills) | cfe5d7b | CC BY-SA 4.0 |
| playwright-skill | [lackeyjb/playwright-skill](https://github.com/lackeyjb/playwright-skill) | bb7e920 | MIT |
| excalidraw-skill | [yctimlin/mcp_excalidraw](https://github.com/yctimlin/mcp_excalidraw) | 7564424 | MIT |
| health | [tw93/claude-health](https://github.com/tw93/claude-health) | b70c0e7 | MIT |
| humanizer-ja | [gonta223/humanizer-ja](https://github.com/gonta223/humanizer-ja) | a1e3436 | MIT |

## 導入時の判断メモ

- **superpowers**: 全14スキルを導入。本来はプラグインだが、各スキルは自己完結したmdなのでプロジェクトスキルとして動作する(セッション開始時の自動注入フックは無し。関連タスクで自動発動する)。
- **planning-with-files**: 英語版のみ導入(ar/de/es/zh/zht版は省略)。テンプレート・スクリプトはスキル内に同梱済み。
- **Understand-Anything**: コアの `understand` のみ導入。ダッシュボード表示などプラグイン専用機能は動かないが、コードベース解析・ナレッジグラフ生成はスキル内スクリプトで完結する。残り8スキル(understand-onboard等)は必要になったら追加。
- **trailofbits/skills**: 39プラグイン中、コード監査以外でも使える汎用3つのみ導入(質問で仕様を確認 / 第二意見 / 危険なAPIパターン集)。本格的なセキュリティ監査系は必要になったら追加(この環境には組み込みの /security-review もある)。
- **claude-health**: メインの `health` スキルのみ導入(check/hunt/read/write等の姉妹スキルは名前が汎用的で衝突しやすいため省略)。
- **導入見送り**: gogcli(steipete/gogcli)はスキルではなくGo製CLIで、Google認証のセットアップが必要。この環境にはGmail / Google Calendar / Google Driveのコネクタが既に接続済みのため不要と判断。
