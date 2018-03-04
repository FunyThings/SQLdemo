--9.1.1利润表-分公司
--增加了当期在职人数的数据获取方法，增加了计算当期人均产值的字段
declare @result table
(公司名称ID Varchar(50),
 公司名称 varchar(max),
 当期服务产值 money,
 当期药械销售 money,
 当期开票 money,
 当期回款 money,
 当期现金支出 money,
 当期出库 money,
 当期在职人数 int
 )
 declare @result1 table
 (公司名称ID int,
  当期项目总金额 money,
  当期开票总金额 money,
  当期回款总金额 money
 )
declare @xmb table
(客户ID int,
 项目金额 money,
 日均产值 money,
 该项目在当期天数 int,
 该项目在该月开票金额 money,
 该项目在该月回款金额 money,
 合同结束日期 date,
 合同开始日期 date,
 项目ID varchar(max),
 合同ID varchar(max),
 执行公司ID int,
 执行公司 varchar(50)
)

declare @kphk table
(客户ID int,
 该项目在该月开票金额 money,
 该项目在该月回款金额 money,
 项目ID varchar(max),
 合同ID varchar(max),
 执行公司ID int,
 执行公司 varchar(50)
)
declare @yx table
(单据号	varchar(max),
 客户名称 varchar(max),
 药械总额 money,
 开票金额 money,
 出库日期 date,
 回款金额 money,
 回款日期 date,
 公司 nvarchar(100),
 公司ID int
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
--set @currenttime = '2017-6-01'
--set @nexttime = '2017-07-01'

--变量申明结束，以下是时间框架
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


insert @xmb--查询档期项目金额
select 
Chose42Key as '客户ID',
Price3 as '项目金额',
(Price3/case when (DATEDIFF(DAY,UserTimeFrom3,UserTimeEnd6))<=0 then 1 else (DATEDIFF(DAY,UserTimeFrom3,UserTimeEnd6)) end) as '日均产值',
(case 
		when UserTimeFrom3<=@currenttime and UserTimeEnd6>=@nexttime then DATEDIFF(DAY,@currenttime,@nexttime)
		when UserTimeFrom3<=@currenttime and UserTimeEnd6<=@nexttime and UserTimeEnd6<=@nexttime then DATEDIFF(DAY,@currenttime,UserTimeEnd6)
		when UserTimeFrom3>=@currenttime and UserTimeFrom3<@nexttime and UserTimeEnd6>=@nexttime then DATEDIFF(DAY,UserTimeFrom3,@nexttime)
		when UserTimeFrom3>=@currenttime and UserTimeFrom3<@nexttime and UserTimeEnd6<=@nexttime then DATEDIFF(DAY,UserTimeFrom3,UserTimeEnd6)
		when UserTimeFrom3>=@currenttime and UserTimeFrom3>=@nexttime then 0
end) as '该项目在当期天数',
		
 ((select SUM(Count10) from dbo.FPMX_OAMainFactBill where Name20=a.Support1 and StartTimeFrom>=@currenttime and StartTimeFrom<@nexttime)*(case when count2=0 then(Count1) else (Count1*Count2) end)) as '该项目在该月开票金额',
 
 (select SUM(Amount20) from dbo.YHFK_OAMainFactBill where Script1=a.Support1 and StartTimeFrom>=@currenttime and StartTimeFrom<@nexttime)*(case when count2=0 then(Count1) else (Count1*Count2) end) as '该项目在该月回款金额',
 UserTimeEnd6 as '合同结束日期',
 UserTimeFrom3 as '合同开始日期',
 FlowFullOrderNum as '项目ID',
 Support1 as '合同ID',
 case when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key) in ('A4','B4')  then (select FatherKey From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key)
	  when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key) in ('A3','B3')  then Chose6Key end as '执行公司ID',
  case when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key) in ('A4','B4')  then (select FatherKeyValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key)
	  when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key) in ('A3','B3')  then Chose6KeyValue end as '执行公司'
from dbo.XMBnew_OAMainFactBill as a
where UserTimeEnd6>=@currenttime

insert @kphk--查询当期开票回款金额
select 

Chose42Key as '客户ID',
 ((select SUM(Count10) from dbo.FPMX_OAMainFactBill where Name20=a.Support1 and StartTimeFrom>=@currenttime and StartTimeFrom<@nexttime)*(case when count2=0 then(Count1) else (Count1*Count2) end)) as '该项目在该月开票金额',
 
 (select SUM(Amount20) from dbo.YHFK_OAMainFactBill where Script1=a.Support1 and StartTimeFrom>=@currenttime and StartTimeFrom<@nexttime)*(case when count2=0 then(Count1) else (Count1*Count2) end) as '该项目在该月回款金额',
 
 FlowFullOrderNum as '项目ID',
 Support1 as '合同ID',
 case when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key) in ('A4','B4')  then (select FatherKey From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key)
	  when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key) in ('A3','B3')  then Chose6Key end as '执行公司ID',
  case when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key) in ('A4','B4')  then (select FatherKeyValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key)
	  when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose6Key) in ('A3','B3')  then Chose6KeyValue end as '执行公司'
from dbo.XMBnew_OAMainFactBill as a

insert @yx--查询当期药械销售金额
select
' <span style=''cursor:pointer;text-decoration:underline;'' FullOrderNum='''+flowfullordernum+''' onclick=''javascript:return oshowcellfullordernum.apply(this);''>'+flowfullordernum+'</span>'as '单据号'	,
Chose1KeyValue as '客户名称', 
Amount1 as '药械总额',
--从收款单中查询该客户该药械销售单的开票金额
isnull((select sum(Amount20) from dbo.FPD_OAMainFactBill where  Name20=a.FlowFullOrderNum),0) as '开票金额',
FlowFinishTime as '出库日期',
--从回款明细中查找该客户的该合同的回款金额
isnull((select sum(Amount20) from dbo.YHFK_OAMainFactBill where  Script1=a.FlowFullOrderNum),0)  as '回款金额',
--从回款明细中查找该客户的该合同回款的金额
(select top 1 StartTimeFrom from dbo.YHFK_OAMainFactBill where  Script1=a.FlowFullOrderNum order by StartTimeFrom desc) as '回款日期',
 case when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose19Key) in ('A4','B4')  then (select FatherKeyValue From dbo.OLAPAgentDim where OLAPKey=a.Chose19Key)
	  when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose19Key) in ('A3','B3')  then Chose19KeyValue end as '公司',
  case when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose19Key) in ('A4','B4')  then (select FatherKey From dbo.OLAPAgentDim where OLAPKey=a.Chose19Key)
	  when (select TagValue From dbo.OLAPAgentDim where OLAPKey=a.Chose19Key) in ('A3','B3')  then Chose19Key end as '公司ID'
from dbo.YXXS_OAMainFactBill as a
where Passed=1 

insert @result1
select 
执行公司ID as '公司名称ID',
SUM(该项目在当期天数*日均产值)  as '当期项目总金额',
(select sum(该项目在该月开票金额) from @kphk where 执行公司ID=b.执行公司ID)  as '当期开票总金额',
(select sum(该项目在该月回款金额) from @kphk where 执行公司ID=b.执行公司ID) as '当期回款总金额'
from @xmb as b
group by 执行公司ID

insert @result
select
olapkey as '公司名称ID', 
MainDemoName as '公司名称', 
isnull((select sum(当期项目总金额) from @result1  where 公司名称ID=b.OLAPKey),0) as '当期产值',
isnull((select sum(回款金额) from @yx as d where 公司ID=b.OLAPKey and 回款日期>=@currenttime and 回款日期<@nexttime),0)as '当期药械销售',
isnull((select sum(当期开票总金额) from @result1  where 公司名称ID=b.OLAPKey),0) as '当期开票',
isnull((select sum(当期回款总金额) from @result1  where 公司名称ID=b.OLAPKey),0) as '当期回款',
isnull((select SUM(Amount20) from dbo.FYDFactBill as fy where (Chose8Key=b.OLAPKey or(select fatherkey from dbo.olapagentdim where olapkey = fy.chose8key)=b.olapkey) and FlowFinishTime>=@currenttime and FlowFinishTime<@nexttime and chose1key=1),0) as '当期现金支出',
isnull((select SUM(Price2) from dbo.CKDFactBill as e  where (chose13key=b.OLAPKey or (select FatherKey from dbo.OLAPAgentDim where OLAPKey=e.Chose13Key)=b.OLAPKey) and FlowFinishTime>=@currenttime and FlowFinishTime<@nexttime),0) as '当期出库',
isnull((select COUNT(OLAPKey) from dbo.OLAPAgentDim as g where IfLeaf=1 and (g.chose2key=b.olapkey or (select FatherKey from dbo.OLAPAgentDim where OLAPKey=g.Chose2Key)=b.OLAPKey) and (IfDel=0 or (IfDel=1 and EndTime>=@currenttime and EndTime<@nexttime))),1) as '当期在职人数'
from dbo.OLAPAgentDim as b
where IfDel=0  and IfLeaf=0 and TagValue='A3'
set @d=@d-1
end
select 
' <span style=''cursor:pointer;text-decoration:underline;'' CurrentReportID ="9_1_1" ReportID="9_1_1_1" Para="公司名称ID" ParaValue="'+公司名称ID+'" onclick=''javascript:return  oshowcellreport.apply(this);''>'+convert(varchar(20),公司名称,1)+'</span>'    as '公司名称', 
convert(nvarchar(20),convert(decimal(15,2),当期服务产值,1),1) as '当期服务产值',
convert(nvarchar(20),convert(decimal(15,2),当期药械销售,1),1) as '当期药械销售',
convert(nvarchar(20),convert(decimal(15,2),当期开票,1),1) as /*'当期开票',*/'当期服务',
convert(nvarchar(20),convert(decimal(15,2),当期回款,1),1) as '当期回款',
convert(nvarchar(20),convert(decimal(15,2),当期现金支出,1),1) as '当期现金支出',
convert(nvarchar(20),convert(decimal(15,2),当期出库,1),1) as '当期出库',
当期在职人数,
--convert( nvarchar(100),(convert(decimal(15,2),当期服务产值,1)/(select case when 当期在职人数 = 0 then 1 else 当期在职人数 end )),1) as '当期人均产值'
 convert(decimal(15,2),(当期服务产值/(select case when 当期在职人数 = 0 then 1 else 当期在职人数 end )),1) as '当期人均产值',
 convert(nvarchar(20),(convert(decimal(15,2),(当期服务产值+当期药械销售-当期现金支出-当期出库),1)),1) as '当期毛利润',
 convert(nvarchar(20),(convert(decimal(15,2),((当期服务产值+当期药械销售-当期现金支出-当期出库)*100/case when isnull((当期服务产值+当期药械销售),1)=0 then 1 else (当期服务产值+当期药械销售)end),1)),1) as '当期毛利率(%)',
 convert(nvarchar(20),(convert(decimal(15,2),(当期回款-当期现金支出-当期出库),1)),1) as '当期净利润',
 convert(nvarchar(20),(convert(decimal(15,2),((当期回款-当期现金支出-当期出库)*100/(case when  isnull(当期回款,1)=0 then 1 else 当期回款 end)),1)),1) as '当期净利率(%)'
 from @result
union all 
select '合计','#','#','#','#','#','#',null,null,'#',null,'#',null
