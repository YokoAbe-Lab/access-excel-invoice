# commit_readme.ps1
# READMEに画像＆PDFリンクを追加した変更をコミット＆プッシュするスクリプト
# リポジトリの場所やブランチ名は必要に応じて変更してください

param(
  [string]$RepoPath = "C:\Users\yo387\Documents\access-excel-invoice",
  [string]$Branch   = ""
)

function Ensure-Git {
  $git = (Get-Command git -ErrorAction SilentlyContinue)
  if (-not $git) { throw "git が見つかりません。Git for Windows をインストールしてください。" }
}

function Detect-Branch {
  if ($Branch -ne "") { return $Branch }
  try {
    $b = (git rev-parse --abbrev-ref HEAD).Trim()
    if ($b) { return $b } else { return "main" }
  } catch { return "main" }
}

try {
  Ensure-Git
  if (-not (Test-Path $RepoPath)) { throw "リポジトリが見つかりません: $RepoPath" }

  Set-Location $RepoPath

  $branchName = Detect-Branch
  Write-Host "ブランチ: $branchName" -ForegroundColor Cyan

  # 変更を最新に（安全のため）
  git pull --rebase origin $branchName

  # 追加する想定ファイル
  $targets = @(
    "README.md",
    "docs\images\invoice_readme_sample.png",
    "docs\invoice_sample_20250814.pdf"
  )

  # 実在するものだけ add
  $existing = @()
  foreach ($t in $targets) {
    if (Test-Path $t) {
      $existing += $t
    } else {
      Write-Host "見つからないためスキップ: $t" -ForegroundColor Yellow
    }
  }

  if ($existing.Count -eq 0) {
    Write-Host "ステージする対象がありません（READMEや画像・PDFのパスをご確認ください）。" -ForegroundColor Yellow
    exit 0
  }

  git add $existing

  # 変更有無チェック
  $status = git status --porcelain
  if (-not $status) {
    Write-Host "コミット対象の変更がありません。" -ForegroundColor Yellow
    exit 0
  }

  $msg = "docs(README): 請求書画像とPDFリンクを追加"
  git commit -m $msg
  git push origin $branchName

  Write-Host "✅ コミット＆プッシュ完了！" -ForegroundColor Green
  Write-Host "GitHubでREADMEを開いて表示をご確認ください。"
}
catch {
  Write-Host "❌ エラー: $($_.Exception.Message)" -ForegroundColor Red
  exit 1
}
# Access→Excel 請求書ツール

Access から「売上台帳」「請求書対象リスト」を出力し、Excel マクロでフィルタ・整形・請求書/PDF を自動作成するツール。

## 構成

## ライセンス
本プロジェクトは [MIT License](./LICENSE) です。

## 変更履歴
詳細は [CHANGELOG.md](./CHANGELOG.md) を参照してください。

[🌐 ACCESS販売管理システム × Excel帳票出力連携プロジェクト（Notion）](https://wide-motion-2bc.notion.site/ACCESS-Excel-24ff5bb7aaa280f59b72d7081825b876?pvs=4)

---

## 🖼 画面サンプル
![請求書サンプル](docs/images/invoice_readme_sample.png)

---

## 📄 請求書PDFサンプル
[📥 請求書サンプルPDFをダウンロード](docs/invoice_sample_20250814.pdf)
