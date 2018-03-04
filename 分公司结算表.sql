--8.5.1各分公司结算表
declare @t1 table
(
 公司ID nvarchar(50),
 公司名称 nvarchar(100),
 本月收入代收 money,
 本月税点及学院费用 money,
 本月代付 money,
 本月转仓 money,
 本月出库 money
)
declare @t table
(
 公司ID nvarchar(50),
 公司名称 nvarchar(100),
 本月收入代收 money,
 本月税点及学院费用 money,
 本月代付 money,
 本月转仓 money,
 本月出库 money,
 结算金额 money
)

declare @tb1 table
(
项目单号 varchar(max),
合同单号 varchar(max),
客户名称ID int,
客户名称 varchar(max),
开票抬头 varchar(max) ,
回款金额 DECIMAL(20,5),
回款银行 nvarchar(100),
回款银行ID nvarchar(50),
合同额分配比例 DECIMAL(20,15),
项目金额分配比例 DECIMAL(20,15),
执行公司ID int,
执行公司 varchar(max),
项目回款金额 DECIMAL(20,5)
)
declare @tb2 table
(
收款单号 varchar(50),
合同编号 varchar(max) ,
客户名称 varchar(max) ,
开票抬头 varchar(max) ,
开票公司ID int,
开票公司 varchar(max),
收款批次 varchar(max),
发票号码 varchar(max),
开票备注 varchar(max),
开票金额 money,
--税点 real,
开票日期 datetime,
开票人 nvarchar(50),
业务员 nvarchar(50),
业务员部门 nvarchar(50),
业务员公司 nvarchar(50),
业务员公司ID nvarchar(50)
)

declare @searchtime varchar(50)
declare @searchtimetype varchar(50)
declare @currenttime datetime
declare @nexttime datetime
declare @d int
declare @dCount int
declare @year int
declare @quater int
declare @month int
declare @day int
set @searchtime='[CurrentTime]'
set @searchtimetype='[TimeType]'
declare @showtime varchar(50)

if @searchtimetype='Y'
begin
set @d=0
end
else if @searchtimetype='Q'
begin
set @d=0
end
else if @searchtimetype='YM'
begin
set @d=0
end
else if @searchtimetype='YMD'
begin
set @d=0
end

	while (@d>=0)
	begin
	 if @searchtimetype='Y'
	 begin
        --循环，如果@d=0,则@currenttime为当前时间，否则在每个循环中往前退一个查询周期，从而形成@d个时间点上的数据曲线
        --分别按年、季、月、日处理，时间框架已经建立好，一般不用修改
		set @currenttime=dateadd(year,-@d,convert(datetime,@searchtime))
		set @nexttime=dateadd(year,1,convert(datetime,@currenttime))
		set @year=year(@currenttime)
        set @showtime=SUBSTRING(convert(varchar(4),@year), 3, 2)+'年'
	 end
	 else if @searchtimetype='Q'
	 begin
		set @currenttime=dateadd(qq,-@d,convert(datetime,@searchtime))
		set @nexttime=dateadd(qq,1,convert(datetime,@currenttime))
		set @year=year(@currenttime)
		set @month=month(@currenttime)
		if @month=1 or @month=2 or @month=3
		begin
		 set @quater=1
		end 
		if @month=4 or @month=5 or @month=6
		begin
		 set @quater=2
		end 
		if @month=7 or @month=8 or @month=9
		begin
		 set @quater=3
		end 
		if @month=10 or @month=11 or @month=12
		begin
		 set @quater=4
		end 

	   if @d=4 or @quater=1
       begin
		set @showtime=SUBSTRING(convert(varchar(4),@year), 3, 2)+'年-'+convert(varchar(1),@quater)+'季'
       end
	   else
       begin
		set @showtime=convert(varchar(1),@quater)+'季'
       end

	 end
	 else if @searchtimetype='YM'
	 begin
		set @currenttime=dateadd(month,-@d,convert(datetime,@searchtime))
		set @nexttime=dateadd(month,1,convert(datetime,@currenttime))
		set @year=year(@currenttime)
		set @month=month(@currenttime)

	   if @d=6 or @month=1
       begin
		set @showtime=SUBSTRING(convert(varchar(4),@year), 3, 2)+'年-'+convert(varchar(2),@month)+'月'
       end
	   else
       begin
		set @showtime=convert(varchar(2),@month)+'月'
       end

	 end
	 else if @searchtimetype='YMD'
	 begin
		set @currenttime='[CurrentTime]'
		set @nexttime='[NextTime]'
		set @year=year(@currenttime)
		set @month=month(@currenttime)
		set @day=day(@currenttime)
  
       if @d=7 or @day=1
       begin
        set @showtime=SUBSTRING(convert(varchar(4),@year), 3, 2)+'-'+convert(varchar(2),@month)+'-'+convert(varchar(2),@day)
       end
	   else
       begin
		set @showtime=convert(varchar(2),@day)
       end
	 end
	 
insert @tb1
select 
x.flowfullordernum as '项目单号',
--' <span style=''cursor:pointer;text-decoration:underline;'' FullOrderNum='''+x.flowfullordernum+''' onclick=''javascript:return oshowcellfullordernum.apply(this);''>'+x.flowfullordernum+'</span>'as '项目单号',
x.Support1 as '合同单号',
--' <span style=''cursor:pointer;text-decoration:underline;'' FullOrderNum='''+x.Support1+''' onclick=''javascript:return oshowcellfullordernum.apply(this);''>'+x.Support1+'</span>'as '合同单号',
x.Chose42Key as '客户名称ID',
x.Chose42KeyValue as '客户名称',
x.Name4 as '开票抬头',
isnull(SUM(h.amount20),0) as '回款金额',
h.Chose4keyvalue as '回款银行',
h.Chose4Key as '回款银行ID',
x.Count1 as '合同额分配比例',
x.Count2 as '项目金额分配比例',
--case when (select tagvalue from dbo.OLAPAgentDim where olapkey = x.Chose6Key)= 4 then x.Chose6Key else (select fatherkey from dbo.OLAPAgentDim where olapkey =x.Chose6Key) end	as	'执行公司ID'	,
--case when (select tagvalue from dbo.OLAPAgentDim where olapkey = x.Chose6Key)= 4 then x.Chose6Keyvalue else (select fatherkeyvalue from dbo.OLAPAgentDim where olapkey =x.Chose6Key) end	as	'执行公司'	,
x.chose6key as '执行公司ID',
x.chose6keyvalue as  '执行公司',
case when x.Count2 = 0 then ((isnull(SUM(h.amount20),0))*x.count1) when x.count1 = 0 then ((isnull(SUM(h.amount20),0))*x.count2)else ((isnull(SUM(h.amount20),0))*x.count1*x.count2)  end as '项目回款金额'
from XMBnew_OAMainFactBill as x inner join dbo.YHFK_OAMainFactBill as h
on x.Support1 = h.script1
where h.StartTimeFrom >=@currenttime and h.StartTimeFrom <@nexttime and h.Chose4Key=[BankLeafKey]
group by x.support1,x.flowfullordernum,x.Chose42Key,x.Chose42KeyValue,x.Name4,x.Count1,x.Count2,x.Chose6Key,x.Chose6Keyvalue,h.Chose4KeyValue,h.Chose4Key
order by x.FlowFullOrderNum

insert @tb2
select
Name2	as	'收款单号'	,
Name20	as	'合同编号'	,
Chose17Keyvalue	as	'客户名称'	,
Name3	as	'开票抬头'	,
Chose1Key	as	'开票公司ID'	,
Chose1Keyvalue	as	'开票公司'	,
Chose4Keyvalue	as	'收款批次'	,
Support20	as	'发票号码'	,
Support1	as	'开票备注'	,
Count10	as	'开票金额'	,
--case when Chose1Key IN( ) then 0.11 else 0.09 end as '税点'
StartTimeFrom	as	'开票日期',
AgentKeyValue as '开票人',
(select Chose15KeyValue from FPD_OAMainFactBill where Name1=f.Name2) as '业务员',
(select Chose1KeyValue from FPD_OAMainFactBill where Name1=f.Name2) as '业务员部门',
(select Chose3KeyValue from FPD_OAMainFactBill where Name1=f.Name2) as '业务员公司',
(select Chose3Key from FPD_OAMainFactBill where Name1=f.Name2) as '业务员公司ID'
from dbo.fpmx_oamainfactbill  as f
where StartTimeFrom>=@currenttime and StartTimeFrom<@nexttime 

insert @t1
select 
OLAPKey as '公司ID',
MainDemoName as '公司名称',
isnull((select SUM(项目回款金额) from @tb1 where 执行公司ID=a.OLAPKey),0) as '本月收入代收',
isnull((select SUM (Amount20) from FYDFactBill where Chose8Key=a.OLAPKey and Chose2Key=[BankLeafKey] and FlowFinishTime>@currenttime and FlowFinishTime<@nexttime),0) as '本月代付',
isnull((select SUM(开票金额)*0.09 from @tb2 where 业务员公司ID=a.OLAPKey and 开票公司ID=[GsxxLeafKey]),0) as '本月税点及学院费用',
isnull((select SUM(Count2) from ZCDFactBill where Chose13Key=a.OLAPKey and Chose17Key=[StoreLeafKey] and FlowFinishTime>@currenttime and FlowFinishTime<@nexttime),0) as '本月转仓',
isnull((select SUM(Price2) from CKDFactBill where Chose13Key=a.OLAPKey and Chose17Key=[StoreLeafKey] and FlowFinishTime>@currenttime and FlowFinishTime<@nexttime),0) as '本月出库'
from dbo.OLAPAgentDim as a
where IfDel=0 and IfLeaf=0 and TagValue in ('A3','B3','A4','B4')
set @d=@d-1
end
insert @t
select 
公司ID as '公司ID',
公司名称 as '公司名称',
本月收入代收 as '本月收入代收',
本月代付 as '本月代付',
本月税点及学院费用 as '本月税点及学院费用',
本月转仓 as '本月转仓',
本月出库 as '本月出库',
(本月收入代收-本月代付-本月转仓-本月出库-本月税点及学院费用) as '结算金额'
from @t1

select 
公司名称,
--本月收入代收,
' <span style=''cursor:pointer;text-decoration:underline;'' CurrentReportID ="8_5_1" ReportID="8_5_1_1" Para="公司ID" ParaValue="'+公司ID+'" onclick=''javascript:return  oshowcellreport.apply(this);''>'+convert(varchar(20),本月收入代收,1)+'</span>'    as '本月收入代收' ,

--本月代付,
' <span style=''cursor:pointer;text-decoration:underline;'' CurrentReportID ="8_5_1" ReportID="8_5_1_4" Para="公司ID" ParaValue="'+公司ID+'" onclick=''javascript:return  oshowcellreport.apply(this);''>'+convert(varchar(20),本月代付,1)+'</span>'    as '本月代付' ,
--本月税点及学院费用,
' <span style=''cursor:pointer;text-decoration:underline;'' CurrentReportID ="8_5_1" ReportID="8_5_1_5" Para="公司ID" ParaValue="'+公司ID+'" onclick=''javascript:return  oshowcellreport.apply(this);''>'+convert(varchar(20),本月税点及学院费用,1)+'</span>'    as '本月税点及学院费用' ,
--本月转仓,
' <span style=''cursor:pointer;text-decoration:underline;'' CurrentReportID ="8_5_1" ReportID="8_5_1_2" Para="公司ID" ParaValue="'+公司ID+'" onclick=''javascript:return  oshowcellreport.apply(this);''>'+convert(varchar(20),本月转仓,1)+'</span>'    as '本月转仓' ,
--本月出库,
' <span style=''cursor:pointer;text-decoration:underline;'' CurrentReportID ="8_5_1" ReportID="8_5_1_3" Para="公司ID" ParaValue="'+公司ID+'" onclick=''javascript:return  oshowcellreport.apply(this);''>'+convert(varchar(20),本月出库,1)+'</span>'    as '本月出库' ,
结算金额
 from @t
where (本月收入代收+本月代付+本月转仓+本月出库+结算金额 )!=0
