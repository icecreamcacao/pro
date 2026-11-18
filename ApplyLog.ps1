param([switch]$Check, [switch]$CheckUI)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$dataDir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ApplyLog'
$dataPath = Join-Path $dataDir 'applications.json'
$statuses = @('관심', '지원 완료', '면접', '합격', '종료')

function Read-Applications([string]$Path) {
    if (Test-Path -LiteralPath $Path) {
        $records = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -ne $records) { $records }
    }
}

function Write-Applications([string]$Path, [object[]]$Records) {
    $directory = Split-Path -Parent $Path
    [void][IO.Directory]::CreateDirectory($directory)
    $temp = Join-Path $directory ([guid]::NewGuid().ToString() + '.tmp')
    try {
        $json = ConvertTo-Json -InputObject $Records -Depth 4
        [IO.File]::WriteAllText($temp, $json, [Text.UTF8Encoding]::new($false))
        if ([IO.File]::Exists($Path)) { [IO.File]::Replace($temp, $Path, [NullString]::Value) }
        else { [IO.File]::Move($temp, $Path) }
    } finally {
        if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    }
}

function Select-Applications([object[]]$Records, [string]$Term, [string]$Status) {
    $Records | Where-Object {
        ($Status -eq '전체' -or $_.status -eq $Status) -and
        (("{0} {1} {2}" -f $_.company, $_.role, $_.notes).IndexOf($Term.Trim(), [StringComparison]::OrdinalIgnoreCase) -ge 0)
    } | Sort-Object -Property @{Expression = { if ($_.deadline) { $_.deadline } else { '9999-99-99' } }}, company
}

if ($Check) {
    $path = Join-Path ([IO.Path]::GetTempPath()) ("applylog-" + [guid]::NewGuid() + ".json")
    try {
        if (@(Read-Applications $path).Count -ne 0) { throw 'Empty read failed' }
        $a = [pscustomobject]@{ company='가상 기업'; role='Developer'; notes='테스트'; status='면접'; deadline='2026-12-01' }
        $b = [pscustomobject]@{ company='Example'; role='QA'; notes=''; status='관심'; deadline='' }
        Write-Applications $path @($b, $a)
        $saved = @(Read-Applications $path)
        if ($saved.Count -ne 2 -or $saved[1].company -cne $a.company) { throw 'UTF-8 round trip failed' }
        if (@(Select-Applications $saved ' developer ' '면접').Count -ne 1) { throw 'Filter failed' }
        if (@(Select-Applications $saved '' '전체')[0].company -ne $a.company) { throw 'Deadline sort failed' }
        $a.status = '합격'
        Write-Applications $path @($a)
        if (@(Read-Applications $path)[0].status -ne '합격') { throw 'Update failed' }
        Write-Applications $path @()
        if (@(Read-Applications $path).Count -ne 0) { throw 'Delete failed' }
        Write-Output 'ApplyLog checks passed.'
    } finally { if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) } }
    return
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$mutex = [Threading.Mutex]::new($false, 'Local\ApplyLogDesktop')
if (-not $mutex.WaitOne(0)) {
    $mutex.Dispose()
    [void][Windows.Forms.MessageBox]::Show('이미 실행 중입니다.', 'ApplyLog')
    return
}

try {
$script:records = if ($CheckUI) { ,@() } else { ,@(Read-Applications $dataPath) }
$form = [Windows.Forms.Form]::new()
$form.Text = 'ApplyLog · 지원 기록'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = [Drawing.Size]::new(1000, 585)
$form.MinimumSize = [Drawing.Size]::new(1016, 624)
$form.Font = [Drawing.Font]::new('맑은 고딕', 10)
$form.BackColor = [Drawing.ColorTranslator]::FromHtml('#F1F5F9')

function Add-Control([string]$Type, [string]$Text, [int]$X, [int]$Y, [int]$Width, [int]$Height) {
    $control = New-Object "Windows.Forms.$Type"
    $control.Text = $Text
    $control.SetBounds($X, $Y, $Width, $Height)
    $form.Controls.Add($control)
    $control
}

$title = Add-Control Label '지원 기록' 24 18 400 43
$title.Font = [Drawing.Font]::new('맑은 고딕', 22, [Drawing.FontStyle]::Bold)
$title.ForeColor = [Drawing.ColorTranslator]::FromHtml('#172554')
$summary = Add-Control Label '' 26 68 620 26
[void](Add-Control Label '회사 · 직무 · 메모 검색' 26 110 210 24)
$search = Add-Control TextBox '' 26 137 420 30
[void](Add-Control Label '지원 상태' 463 110 150 24)
$filter = Add-Control ComboBox '' 463 137 155 30
$filter.DropDownStyle = 'DropDownList'
$filter.Items.AddRange([object[]](@('전체') + $statuses))
$filter.SelectedIndex = 0

$grid = Add-Control DataGridView '' 26 182 592 375
$grid.Anchor = 'Top, Bottom, Left, Right'
$grid.ReadOnly = $true
$grid.MultiSelect = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [Drawing.Color]::White
$grid.RowTemplate.Height = 32
foreach ($column in @('회사', '직무', '상태', '마감일')) {
    $index = $grid.Columns.Add($column, $column)
    $grid.Columns[$index].SortMode = 'NotSortable'
}

$editor = @{}
[void](Add-Control Label '지원 정보' 645 104 320 28)
foreach ($field in @(
    @{key='company'; label='회사 (필수)'; y=140},
    @{key='role'; label='직무 (필수)'; y=201},
    @{key='status'; label='상태'; y=262},
    @{key='deadline'; label='마감일 (YYYY-MM-DD, 선택)'; y=323},
    @{key='notes'; label='메모 · 다음 할 일'; y=384}
)) {
    $label = Add-Control Label $field.label 645 $field.y 324 24
    $label.Anchor = 'Top, Right'
    $kind = if ($field.key -eq 'status') { 'ComboBox' } else { 'TextBox' }
    $height = if ($field.key -eq 'notes') { 85 } else { 30 }
    $input = Add-Control $kind '' 645 ($field.y + 24) 328 $height
    $input.Anchor = 'Top, Right'
    if ($field.key -eq 'notes') {
        $input.Multiline = $true
        $input.ScrollBars = 'Vertical'
    }
    if ($field.key -eq 'status') {
        $input.DropDownStyle = 'DropDownList'
        $input.Items.AddRange([object[]]$statuses)
    }
    $editor[$field.key] = $input
}

$save = Add-Control Button '저장' 645 516 108 36
$reset = Add-Control Button '새 지원' 761 516 102 36
$delete = Add-Control Button '삭제' 871 516 102 36
foreach ($button in @($save, $reset, $delete)) { $button.Anchor = 'Top, Right' }
$save.BackColor = [Drawing.ColorTranslator]::FromHtml('#1D4ED8')
$save.ForeColor = [Drawing.Color]::White

$script:currentId = ''
function Clear-Editor {
    $script:currentId = ''
    $grid.ClearSelection()
    foreach ($key in @('company','role','deadline','notes')) { $editor[$key].Clear() }
    $editor.status.SelectedIndex = 0
    $delete.Enabled = $false
}

function Refresh-Grid {
    $grid.Rows.Clear()
    $visible = @(Select-Applications $script:records $search.Text $filter.Text)
    $today = (Get-Date).Date
    $soon = 0
    foreach ($record in $visible) {
        $row = $grid.Rows.Add([object[]]@($record.company, $record.role, $record.status, $record.deadline))
        $grid.Rows[$row].Tag = $record.id
    }
    foreach ($record in $script:records) {
        if ($record.deadline -and $record.status -eq '관심') {
            $date = [datetime]::ParseExact($record.deadline, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            if ($date -ge $today -and $date -le $today.AddDays(7)) { $soon++ }
        }
    }
    $summary.Text = '전체 {0}건  ·  검색 결과 {1}건  ·  7일 이내 마감 예정 {2}건' -f $script:records.Count, $visible.Count, $soon
    $grid.ClearSelection()
}

function Save-Changes([object[]]$Next) {
    try {
        Write-Applications $dataPath $Next
        $script:records = $Next
        Clear-Editor
        Refresh-Grid
    } catch {
        [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, '저장 오류')
    }
}

$loadRow = {
    param($sender, $event)
    if ($event.RowIndex -lt 0) { return }
    $id = $grid.Rows[$event.RowIndex].Tag
    if (-not $id) { return }
    $record = $script:records | Where-Object id -EQ $id | Select-Object -First 1
    $script:currentId = $record.id
    foreach ($key in $editor.Keys) { $editor[$key].Text = [string]$record.$key }
    $delete.Enabled = $true
}
$grid.Add_CellEnter($loadRow)
$grid.Add_CellClick($loadRow)
$search.Add_TextChanged({ Refresh-Grid })
$filter.Add_SelectedIndexChanged({ Refresh-Grid })
$reset.Add_Click({ Clear-Editor; $editor.company.Focus() | Out-Null })
$save.Add_Click({
    $company = $editor.company.Text.Trim()
    $role = $editor.role.Text.Trim()
    $deadline = $editor.deadline.Text.Trim()
    $date = [datetime]::MinValue
    if (-not $company -or -not $role) {
        [void][Windows.Forms.MessageBox]::Show('회사와 직무를 입력하세요.', '입력 확인')
        return
    }
    if ($deadline -and -not [datetime]::TryParseExact($deadline, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) {
        [void][Windows.Forms.MessageBox]::Show('마감일은 YYYY-MM-DD 형식의 실제 날짜를 입력하세요.', '입력 확인')
        return
    }
    $id = if ($script:currentId) { $script:currentId } else { [guid]::NewGuid().ToString() }
    $record = [pscustomobject]@{ id=$id; company=$company; role=$role; deadline=$deadline; status=$editor.status.Text; notes=$editor.notes.Text.Trim() }
    Save-Changes (@($script:records | Where-Object id -NE $id) + $record)
})
$delete.Add_Click({
    $result = [Windows.Forms.MessageBox]::Show('선택한 지원 기록을 삭제할까요?', '기록 삭제', 'YesNo', 'Question')
    if ($result -eq [Windows.Forms.DialogResult]::Yes) {
        Save-Changes @($script:records | Where-Object id -NE $script:currentId)
    }
})
Refresh-Grid
Clear-Editor
if ($CheckUI) {
    if ($grid.Columns.Count -ne 4 -or $editor.status.Items.Count -ne 5) { throw 'UI setup failed' }
    $form.Dispose()
    Write-Output 'ApplyLog UI initialization passed.'
    return
}
[void]$form.ShowDialog()
} catch {
    if ($CheckUI) { throw }
    [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'ApplyLog 실행 오류')
} finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
