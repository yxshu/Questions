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
		C_N_Id varchar not null,-----�½ڱ��+allid
		SN int  not null,-----����ԭ���
		SNID varchar  not null,-----�½ڱ��+SN
		Subj varchar  null,-----��Ŀ
		Chapter varchar  null,-----�±���
		Node varchar null,-----�ڱ���
		Title nvarchar not null,-----���
		Choosea varchar null,-----ѡ��A
		Chooseb varchar null,-----ѡ��B
		Choosec varchar null,-----ѡ��C
		Choosed varchar null,-----ѡ��D
		Answer int not null check(Answer in(1,2,3,4)),-----�ο���
		Explain varchar null,-----����
		ImageAddress varchar null----ͼƬ��ַ
)
if exists (select*from sysobjects where name='PanduanQuestion')
drop table PanduanQuestion
create table PanduanQuestion----�ж��������
	(
	Id int identity(1,1)primary key,---�Զ����
		AllID int not null,-----����������Զ����
		C_N_Id varchar not null,-----�½ڱ��+allid
		SN int  not null,-----����ԭ���
		SNID varchar  not null,-----�½ڱ��+SN
		Subj varchar  null,-----��Ŀ
		Chapter varchar  null,-----�±���
		Node varchar null,-----�ڱ���
		Title nvarchar not null,-----���
		Answer bit not null,-----�ο���
		Explain varchar null,-----����
		ImageAddress varchar null----ͼƬ��ַ
)