------------------------------------------------ �������ݿⲿ�� ------------------------------------------------


use master -- ���õ�ǰ���ݿ�Ϊmaster,�Ա����sysdatabases��
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
create table ChooseQuestion-----ѡ���������
	(
	Id int identity(1,1)primary key,---�Զ����
		AllID int not null,-----����������Զ����
		C_N_Id varchar(1000) not null,-----�½ڱ��+allid
		SN int  not null,-----����ԭ���
		SNID varchar(1000)  not null,-----�½ڱ��+SN
		Subj varchar(1000)  null,-----��Ŀ
		Chapter varchar(1000)  null,-----�±���
		Node varchar(1000) null,-----�ڱ���
		Title varchar(8000) not null,-----���
		Choosea varchar(8000) null,-----ѡ��A
		Chooseb varchar(8000) null,-----ѡ��B
		Choosec varchar(8000) null,-----ѡ��C
		Choosed varchar(8000) null,-----ѡ��D
		Answer int not null check(Answer in(0,1,2,3,4)),-----�ο���
		Explain varchar(max) null,-----����
		ImageAddress varchar(8000) null,----ͼƬ��ַ
		Remark varchar(8000) null----��ע
)
if exists (select*from sysobjects where name='PanduanQuestion')
drop table PanduanQuestion
create table PanduanQuestion----�ж��������
	(
	Id int identity(1,1)primary key,---�Զ����
		AllID int not null,-----����������Զ����
		C_N_Id varchar(1000) not null,-----�½ڱ��+allid
		SN int  not null,-----����ԭ���
		SNID varchar(1000)  not null,-----�½ڱ��+SN
		Subj varchar(1000)  null,-----��Ŀ
		Chapter varchar(1000)  null,-----�±���
		Node varchar(1000) null,-----�ڱ���
		Title varchar(8000) not null,-----���
		Answer bit not null,-----�ο���
		Explain varchar(max) null,-----����
		ImageAddress varchar(8000) null,----ͼƬ��ַ
		Remark varchar(8000) null----��ע
)