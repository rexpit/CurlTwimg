# 一部のサイトの画像保存を簡単化します。
#
# 基本的な使い方
# - 引数に保存したい画像の URL を指定するだけです。
#
# 主な機能
# - Twitter 画像を orig (オリジナルサイズ) でダウンロードします。
# - pixiv 画像をダウンロードするために、HTTP リクエストに Referer を自動で設定します。
# - URL のファイルをダウンロードする際、"Last-Modified" を更新日時に反映します。

# コマンドライン引数
Param
(
    [string]$URL,
    [string]$OutFile,
    [string]$Referer
)

# 定数
$HostName_twimg = "pbs.twimg.com"
$HostName_pximg = "i.pximg.net"
$Referer_pixiv = "https://www.pixiv.net"

# .NETアセンブリをロード
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 引数指定がなければ対話型実行
function MainInteractive
{
    $inputUrl = InputBox "ダウンロードする URL を入力してください。" "URL 入力"
    if ( -not [string]::IsNullOrEmpty($inputUrl) )
    {
        Write-Host "Input URL: $inputUrl"

        # 入力 URL から変更 URL とデフォルトファイル名（とついでに Referer）を取得
        $outputUrl, $origFileName, $referer = Get_OrigUrl_OutFileName $inputUrl

        # 入力 URL と変更 URL が異なる場合のみ表示
        if ($inputUrl -ne $outputUrl)
        {
            Write-Host "orig URL : $outputUrl"
        }
        Write-Host "FileName : $origFileName"
        # Referer がある場合のみ表示
        if (-not [string]::IsNullOrEmpty($referer))
        {
            Write-Host "Referer  : $referer"
        }

        # SaveFileDialog を定義
        $sfd = [System.Windows.Forms.SaveFileDialog]::new()
        $sfd.Title = "名前を付けて保存"
        $sfd.Filter = "すべてのファイル|*.*|画像ファイル|*.png;*.jpg;*.gif|JPEG|*.jpg|PNG|*.png|GIF|*.gif"
        $sfd.FileName = $origFileName
        # ファイル選択ダイアログを表示
        if ($sfd.ShowDialog() -eq "OK")
        {
            # 1行空ける
            Write-Host
            # ダウンロード実行
            DownloadFileAndSave_Invoke_WebRequest $outputUrl $sfd.FileName $referer
        }
        else
        {
            Write-Host "キャンセルしました。"
        }
    }
    else
    {
        Write-Host "キャンセルしました。"
    }
}

# 引数指定があれば自動実行
# 引数
# - $paramUrl: 入力 URL
# - $paramFileName: 入力 ファイル名
# - $paramReferer: 入力 Referer
function MainAuto([string]$paramUrl, [string]$paramFileName=[string]::Empty, [string]$paramReferer=[string]::Empty)
{
    $inputUrl = $paramUrl
    if ( -not [string]::IsNullOrEmpty($inputUrl) )
    {
        Write-Host "Input URL: $inputUrl"

        # 入力 URL から変更 URL とデフォルトファイル名（とついでに Referer）を取得
        $outputUrl, $origFileName, $referer = Get_OrigUrl_OutFileName $inputUrl

        # 入力 URL と変更 URL が異なる場合のみ表示
        if ($inputUrl -ne $outputUrl)
        {
            Write-Host "orig URL : $outputUrl"
        }
        # ファイル名の引数指定がある場合はそちらを使用
        if ( -not [string]::IsNullOrEmpty($paramFileName) )
        {
            $saveFileName = $paramFileName
        }
        # 引数指定がない場合は URL からファイル名を指定
        else
        {
            Write-Host "FileName : $origFileName"
            $saveFileName = $origFileName
        }
        # Referer の指定があればそれを使用
        if (-not [string]::IsNullOrEmpty($paramReferer))
        {
            $referer = $paramReferer
        }
        # Referer がある場合のみ表示
        if (-not [string]::IsNullOrEmpty($referer))
        {
            Write-Host "Referer  : $referer"
        }

        # 指定のファイル名が存在すれば中止
        if (Test-Path $saveFileName)
        {
            Write-Host ("{0} がすでに存在します。" -f $saveFileName)
        }
        # 指定のファイル名が存在しなければダウンロード開始
        else
        {
            # 1行空ける
            Write-Host
            # ダウンロード実行
            DownloadFileAndSave_Invoke_WebRequest $outputUrl $saveFileName $referer
        }
    }
}

# Twitter 画像 URL から orig 画像 URL とファイル名を取得します。
# 引数
# - $url: Twitter 画像 URL
# 戻り値
# - ($newUrl, $fileName, $referer): 要素3の配列を返します。
#  - $newUrl: orig 画像 URL
#  - $fileName: ファイル名
#  - $referer: リファラー
function Get_OrigUrl_OutFileName([string]$url)
{
    $uri = [System.Uri]::new($url)
    # デフォルトの出力はそのまま
    $newUrl = $url
    $fileName = (Split-Path $uri.AbsolutePath -Leaf)
    $referer = [string]::Empty

    # Twitter 画像の場合
    if ( $uri.Host -eq $Script:HostName_twimg )
    {
        Write-Host "Twitter 画像モード"
        # "?format=拡張子&name=…" があったら "?format=拡張子&name=orig" にする。
        if (-not [string]::IsNullOrEmpty($uri.Query))
        {
            # クエリ文字列を抽出する(ただしすでに "name=" があれば削除し、"name=orig" を付与する)。
            $queryString = "?{0}&name=orig" -f ( ( $uri.Query.TrimStart("?").Split("&") |
                Where-Object {$_.Split("=")[0] -ne "name"} ) -join "&" )
            # クエリ文字列に format= がなければエラー
            if ($queryString -notmatch "[\?&]format=")
            {
                throw "URL に ""format=拡張子"" が見つかりません。"
            }
            $ext = $queryString -replace "^.*[\?&]format=([^&]*).*$","`$1"
            # URL を生成
            $newUrl = ($url -replace "\?.*","") + $queryString
            $fileName = ("{0}.{1}" -f (Split-Path $uri.AbsolutePath -Leaf),$ext)
        }
        # 拡張子があれば、":orig" をつける。
        elseif ( $url -match "\.\w+(?:\:\w+)?$" )
        {
            $newUrl = $url -replace "^(.*\.[^\:/]*)(?:\:.*)?$","`$1:orig"
            $fileName = (Split-Path $uri.AbsolutePath -Leaf) -replace "(?:\:[^\:]*)?$",""
        }
        # それ以外は想定外
        else
        {
            throw "想定外の URL です。"
        }
    }
    # pixiv 画像の場合
    elseif ( $uri.Host -eq $Script:HostName_pximg )
    {
        Write-Host "pixiv 画像モード"
        $referer = $Script:Referer_pixiv
    }
    return $newUrl, $fileName, $referer
}

# URL のファイルをダウンロードして保存します。
# 引数
# - $url: ダウンロードする URL
# - $saveFileName: 保存ファイル名
# - $referer: リファラー
function DownloadFileAndSave_WebClient([string]$url, [string]$saveFileName, [string]$referer=[string]::Empty)
{
    Write-Host "Download $url"
    $client = [System.Net.WebClient]::new()
    $uri = [System.Uri]::new($url)
    # Referer がある場合はヘッダーに追加する
    if ( -not [string]::IsNullOrEmpty($referer) )
    {
        $client.Headers.Add("Referer", $referer)
    }
    $client.DownloadFile($uri, $saveFileName)
    Write-Host "ダウンロードが完了しました。"
    Write-Host ("FileName: {0}" -f $saveFileName)
    # Last-Modified があれば更新日時をその値に変更する。
    if ( $client.ResponseHeaders.Keys -contains "Last-Modified" )
    {
        $lastModifiedStr = $client.ResponseHeaders["Last-Modified"]
        (Get-Item $saveFileName).LastWriteTime = (Get-Date -Date $lastModifiedStr)
        Write-Host "Last-Modified: $lastModifiedStr"
    }
}

# URL のファイルをダウンロードして保存します。
# 引数
# - $url: ダウンロードする URL
# - $saveFileName: 保存ファイル名
# - $referer: リファラー
function DownloadFileAndSave_Invoke_WebRequest([string]$url, [string]$saveFileName, [string]$referer=[string]::Empty)
{
    Write-Host "Download $url"
    $headers = @{}
    if (-not [string]::IsNullOrEmpty($referer) )
    {
        $headers["Referer"] = $referer
    }
    $webRequest = Invoke-WebRequest $url -Headers $headers
    Write-Host "ダウンロードが完了しました。"
    # バイナリ保存
    [System.IO.File]::WriteAllBytes($saveFileName, $webRequest.Content)
    Write-Host ("FileName: {0}" -f $saveFileName)
    # Last-Modified があれば更新日時をその値に変更する。
    if ( $webRequest.Headers.Keys -contains "Last-Modified" )
    {
        # PowerShell バージョンによってデータ型が異なるので確認
        $lastModifiedValue = $webRequest.Headers["Last-Modified"]
        if ( $lastModifiedValue.GetType() -eq [string] )
        {
            $lastModifiedStr = $lastModifiedValue
        }
        elseif ( $lastModifiedValue.GetType().BaseType -eq [array] )
        {
            $lastModifiedStr = $lastModifiedValue[0]
        }
        else
        {
            throw "WebRequest.Headers[""Last-Modified""] のデータ型が想定外です。"
        }
        (Get-Item $saveFileName).LastWriteTime = (Get-Date -Date $lastModifiedStr)
        Write-Host "Last-Modified: $lastModifiedStr"
    }
}

# 文字列入力画面を表示します。
# 引数
# - $prompt: 画面に表示するメッセージ
# - $title: ウィンドウタイトル
# - $default: テキストボックス初期値
function InputBox([string]$prompt, [string]$title = "", [string]$default = "")
{
    # フォームの作成
    $form = [System.Windows.Forms.Form]::new()
    $form.Text = $title
    $form.AutoScaleDimensions = [System.Drawing.SizeF]::new(6, 12)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $form.Size = [System.Drawing.Size]::new(400, 200)
    $form.MinimumSize = [System.Drawing.Size]::new(227, 131)
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # ラベルの設定
    $label = [System.Windows.Forms.Label]::new()
    $label.Anchor = [System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor
            [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    #$label.AutoSize = $true
    $label.Location = [System.Drawing.Point]::new(12, 9)
    $label.Size = [System.Drawing.Size]::new(358, 81)
    $label.TabIndex = 0
    $label.Text = $prompt

    # 入力ボックスの設定
    $textBox = [System.Windows.Forms.TextBox]::new()
    $textBox.Anchor = [System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right)
    $textBox.Location = [System.Drawing.Point]::new(12, 101)
    $textBox.Size = [System.Drawing.Size]::new(358, 19)
    $textBox.TabIndex = 1
    $textBox.Text = $default
    $textBox.Add_KeyDown({
        # コントロールキーを押しながら、Aキーを押した場合
        if ($_.Control -and $_.KeyCode -eq "A") {
            # テキストをすべて選択
            $this.SelectAll()
            # ビープ音が鳴らないように、KeyPressイベントを抑制する
            $_.SuppressKeyPress = $true
        }
    })

    # OKボタンの設定
    $OKButton = [System.Windows.Forms.Button]::new()
    $OKButton.Anchor = [System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $OKButton.Location = [System.Drawing.Point]::new(216, 126)
    $OKButton.Size = [System.Drawing.Size]::new(75, 23)
    $OKButton.TabIndex = 2
    $OKButton.Text = "OK"
    $OKButton.DialogResult = "OK"
    $OKButton.UseVisualStyleBackColor = $true

    # キャンセルボタンの設定
    $CancelButton = [System.Windows.Forms.Button]::new()
    $CancelButton.Anchor = [System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $CancelButton.Location = [System.Drawing.Point]::new(297, 126)
    $CancelButton.Size = [System.Drawing.Size]::new(75, 23)
    $CancelButton.TabIndex = 3
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = "Cancel"
    $CancelButton.UseVisualStyleBackColor = $true

    # キーとボタンの関係
    $form.AcceptButton = $OKButton
    $form.CancelButton = $CancelButton

    # ボタン等をフォームに追加
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($OKButton)
    $form.Controls.Add($CancelButton)

    # フォームを常に手前に表示
    $form.Topmost = $true
    # フォームをアクティブにし、テキストボックスにフォーカスを設定
    $form.Add_Shown({$textBox.Select()})

    # フォームを表示させ、その結果を受け取る
    $dialogResult = $form.ShowDialog()

    # 結果による処理分岐
    if ($dialogResult -eq "OK")
    {
        $result = $textBox.Text
    }
    else
    {
        $result = ""
    }
    return $result
}

# 処理実行
try
{
    # 引数指定がなければダイアログ表示
    if ( [string]::IsNullOrEmpty($URL) )
    {
        # 対話型実行
        MainInteractive
    }
    # 引数があれば自動実行
    else
    {
        # 第1引数を URL とする。
        MainAuto $URL $OutFile $Referer
    }
}
catch
{
    # 例外処理
    Write-Error $_.Exception
    # 引数指定がなければダイアログも表示する。
    if ( $args.Count -le 0 )
    {
        $dialogResult = [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", "OK", "Error")
    }
}
