USE [mybase]
GO

CREATE TABLE [dbo].[atable](
	[aid] [int] IDENTITY(1,1) NOT NULL,
	[adate] [date] NOT NULL,
	[anumber] [bigint] NOT NULL,
	[avalue] [float] NOT NULL,
	[atext] [varchar](250) NULL,
)

GO
 
CREATE TABLE [dbo].[btable](
    [bid] [int] IDENTITY(1,1) NOT NULL,
	[bdate] [date] NOT NULL,
	[bnumber] [bigint] NOT NULL,
	[bvalue] [float] NOT NULL,
	[btext] [varchar](250) NULL,
)

GO

CREATE TABLE [dbo].[ctable](
	[cid] [int] IDENTITY(1,1) NOT NULL,
    [cdate] [date] NOT NULL,
    [cnumber] [bigint] NOT NULL,
    [cvalue] [float] NOT NULL,
    [ctext] [varchar](250) NULL,
)

GO
 
 CREATE TABLE [dbo].[log](
	[id] [bigint] IDENTITY(1,1) NOT NULL,
	[dt] [datetime] NOT NULL,
	[ip] [varchar](50) NOT NULL,
	[browser] [varchar](250) NULL,
	[action] [varchar](50) NULL,
	[value] [varchar](500) NULL
)

GO

declare @i int =0
while @i<30
	begin
		declare @mydate date, @mynumber int, @myvalue float, @mytext varchar(200)
		select @mydate = cast(getdate()+1000*rand() As Date)
		select @mynumber = round(1000*rand(),0)
		select @myvalue =  100*rand()
		declare @j int =0
		declare @s varchar(12) = ''
		while @j<100
			begin
				select @s += char(rand(checksum(newid()))* 25 + 65)
				set @j=@j+1
			end
		select @mytext = right(@s,rand()*100)
		if @i <10
			begin
				insert into [dbo].[atable]([adate],[anumber],[avalue],[atext]) values(@mydate, @mynumber, @myvalue, @mytext)
			end
		else if @i >=10 and @i <20
			begin
				insert into [dbo].[btable]([bdate],[bnumber],[bvalue],[btext]) values(@mydate, @mynumber, @myvalue, @mytext)
			end
		else if @i >=20
			begin
				insert into [dbo].[ctable]([cdate],[cnumber],[cvalue],[ctext]) values(@mydate, @mynumber, @myvalue, @mytext)
			end
		set @i=@i+1
	end