# 脚本：云营业厅业务办理汇报数据下载.ps1
# 作用说明：将云营业厅业务办理情况的数据获取后筛选，并存入文本文件
# 作者：
# 电子邮箱：
# 版本：0.1
# 日期：2014-01-28
# 补充说明：如果系统运行时提示“此系统中禁止执行脚本”，可在powershell中执行set-ExecutionPolicy RemoteSigned命令后执行此脚本

Write-Host "*** 云营业厅业务办理汇报下载 ***`r`n" -foregroundcolor Cyan

# 计算昨天的日期，格式 2014-01-07
$time = "{0:yyyy-MM-dd}" -f (Get-Date).AddDays(-1)
# 允许用户自己录入日期
$timeinput = Read-Host "请输入您要导出记录的时间，格式为 yyyy-MM-dd（留空默认为 $time ）"
If ($timeinput)
{
$time = $timeinput
}

# 创建*两个*记录文件，并向文件中写入记录的所属日期
$streamall = [System.IO.StreamWriter] "$time 汇总业务.txt"
$streamall.WriteLine("日期：$time`r`n")

$streamsingle = [System.IO.StreamWriter] "$time 明细业务.txt"
$streamsingle.WriteLine("日期：$time`r`n")

# 定义营业厅id和名称，准备遍历
$officeids = @('1', '2', '6', '4', '5', '3', '')
$officenames = @('郑汴路营业厅', '碧沙岗营业厅', '天津路营业厅', '凤化街营业厅', '纱厂路营业厅', '涧东路营业厅', '全部')
$p = 0
foreach ($id in $officeids)
{
# 向*明细业务*的记录文件中写入营业厅名称
$streamsingle.WriteLine($officenames[$p])

# 组装请求的网址，并将结果按Json分析
$url = "http://10.88.131.228:9080/cbh_business_service/main/transact/rateList.html?officeid=$id&amp;starttime=$time+00%3A00%3A00&amp;endtime=$time+23%3A59%3A59&amp;charttype=hour"
#echo $url
$content = (Invoke-WebRequest -Uri $url).Content | ConvertFrom-Json

#echo $content.Rows.Count
#echo $content.Rows[0].name $content.Rows[0].status1 $content.Rows[0].status0
# 根据Json的内容，对每个营业厅的每一项业务进行记录。Json中的*统计*一列，不写入*明细业务*记录中
for ($i = 0; $i -le $content.Rows.Count; $i++)
{
$record = $content.Rows[$i]
If ($i -eq $content.Rows.Count - 1)
{
# 将*统计*列的数据 营业厅名称、总笔数、成功笔数、失败笔数 写入*汇总业务*记录中
$streamall.WriteLine([string]::Format("{0}`t{1}`t{2}`t{3}", $officenames[$p], $record.status1 + $record.status0, $record.status1, $record.status0))
}
else
{
# 将除了*统计*列外的每一项的 名称、成功笔数、失败笔数 写入*明细业务*记录中
$streamsingle.WriteLine([string]::Format("{0}`t{1}`t{2}", $record.name, $record.status1, $record.status0))
}
}

$p++
}

# 关闭记录文件
$streamall.Close()
$streamsingle.Close();

Write-Host "`r`n...文件导出成功，按任意键退出。"
$null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")