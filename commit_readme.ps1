# commit_readme.ps1 〔自動追記つき版〕
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

  git pull --rebase origin $branchName

  # ===== README へ自動追記 =====
  $readme = "README.md"
  if (-not (Test-Path $readme)) { throw "README.md が見つかりません。" }

  $block = @"
[🌐 ACCESS販売管理システム × Excel帳票出力連携プロジェクト（Notion）](https://wide-motion-2bc.notion.site/ACCESS-Excel-24ff5bb7aaa280f59b72d7081825b876?pvs=4)

---

## 🖼 画面サンプル

![請求書サンプル](docs/images/invoice_readme_sample.png)

---

## 📄 請求書PDFサンプル

[📥 請求書サンプルPDFをダウンロード](docs/invoice_sample_20250814.pdf)
"@

  $text = Get-Content $readme -Raw
  if ($text -notmatch "請求書サンプルPDFをダウンロード") {
    Add-Content $readme "`r`n$block`r`n"
    Write-Host "README.md にブロックを追記しました。" -ForegroundColor Green
  } else {
    Write-Host "README.md には既に同等のブロックがあります。追記はスキップします。" -ForegroundColor Yellow
  }

  # ステージ対象
  $targets = @(
    "README.md",
    "docs\images\invoice_readme_sample.png",
    "docs\invoice_sample_20250814.pdf"
  )

  $existing = @()
  foreach ($t in $targets) {
    if (Test-Path $t) { $existing += $t } else { Write-Host "見つからないためスキップ: $t" -ForegroundColor Yellow }
  }

  if ($existing.Count -eq 0) { Write-Host "ステージ対象がありません。" -ForegroundColor Yellow; exit 0 }

  git add $existing
  $status = git status --porcelain
  if (-not $status) { Write-Host "コミット対象の変更がありません。" -ForegroundColor Yellow; exit 0 }

  git commit -m "docs(README): Notionリンク・画像・PDFリンクを追加"
  git push origin $branchName

  Write-Host "✅ コミット＆プッシュ完了！ GitHubでREADMEを確認してください。" -ForegroundColor Green
}
catch {
  Write-Host "❌ エラー: $($_.Exception.Message)" -ForegroundColor Red
  exit 1
}
