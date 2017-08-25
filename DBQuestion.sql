------------------------------------------------ 生成数据库部分 ------------------------------------------------


use master -- 设置当前数据库为master,以便访问sysdatabases表
if exists(select * from sysdatabases where name='Question') 
drop database Question      
CREATE DATABASE Question ON  PRIMARY 
(
	 NAME = N'Question', 
	 FILENAME = N'd:\Question.mdf' , 
	 SIZE = 10240KB , 
	 FILEGROWTH = 1024KB 
 )
 LOG ON 
( 
	NAME = N'Question_log', 
	FILENAME = N'd:\Question_log.ldf' , 
	SIZE = 1024KB , 
	FILEGROWTH = 10%
)
GO
use Question
if exists (select*from sysobjects where name='ChooseQuestion')
drop table ChooseQuestion
create table ChooseQuestion-----选择题试题表
	(
	Id int identity(1,1)primary key,---自动编号
		AllID int not null,-----来自试题的自动编号
		C_N_Id varchar(1000) not null,-----章节编号+allid
		SN int  not null,-----试题原编号
		SNID varchar(1000)  not null,-----章节编号+SN
		Subj varchar(1000)  null,-----科目
		Chapter varchar(1000)  null,-----章标题
		Node varchar(1000) null,-----节标题
		Title varchar(8000) not null,-----题干
		Choosea varchar(8000) null,-----选项A
		Chooseb varchar(8000) null,-----选项B
		Choosec varchar(8000) null,-----选项C
		Choosed varchar(8000) null,-----选项D
		Answer int not null check(Answer in(0,1,2,3,4)),-----参考答案
		Explain varchar(max) null,-----解析
		ImageAddress varchar(8000) null,----图片地址
		Remark varchar(8000) null----备注
)
if exists (select*from sysobjects where name='PanduanQuestion')
drop table PanduanQuestion
create table PanduanQuestion----判断题试题表
	(
	Id int identity(1,1)primary key,---自动编号
		AllID int not null,-----来自试题的自动编号
		C_N_Id varchar(1000) not null,-----章节编号+allid
		SN int  not null,-----试题原编号
		SNID varchar(1000)  not null,-----章节编号+SN
		Subj varchar(1000)  null,-----科目
		Chapter varchar(1000)  null,-----章标题
		Node varchar(1000) null,-----节标题
		Title varchar(8000) not null,-----题干
		Answer bit not null,-----参考答案
		Explain varchar(max) null,-----解析
		ImageAddress varchar(8000) null,----图片地址
		Remark varchar(8000) null----备注
)