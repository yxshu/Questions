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
		C_N_Id varchar not null,-----章节编号+allid
		SN int  not null,-----试题原编号
		SNID varchar  not null,-----章节编号+SN
		Subj varchar  null,-----科目
		Chapter varchar  null,-----章标题
		Node varchar null,-----节标题
		Title nvarchar not null,-----题干
		Choosea varchar null,-----选项A
		Chooseb varchar null,-----选项B
		Choosec varchar null,-----选项C
		Choosed varchar null,-----选项D
		Answer int not null check(Answer in(1,2,3,4)),-----参考答案
		Explain varchar null,-----解析
		ImageAddress varchar null----图片地址
)
if exists (select*from sysobjects where name='PanduanQuestion')
drop table PanduanQuestion
create table PanduanQuestion----判断题试题表
	(
	Id int identity(1,1)primary key,---自动编号
		AllID int not null,-----来自试题的自动编号
		C_N_Id varchar not null,-----章节编号+allid
		SN int  not null,-----试题原编号
		SNID varchar  not null,-----章节编号+SN
		Subj varchar  null,-----科目
		Chapter varchar  null,-----章标题
		Node varchar null,-----节标题
		Title nvarchar not null,-----题干
		Answer bit not null,-----参考答案
		Explain varchar null,-----解析
		ImageAddress varchar null----图片地址
)