use VideoVSTO
go 

drop database Georeports;
go

create database Georeports
go

use Georeports

create table CatalogsReports 
(
	Id int identity not null primary key,
	BeginDate date not null, -- начало облета
	EndDate date not null, -- конец облета
	Name nvarchar(256) not null, -- название каталога
	count int not null
);

create table DocumentsReports
(
	Id int identity not null,
	CatalogsId int not null,
	NameOfReportsSet nvarchar(256) not null,
	CreateDate date not null,
	WordDocument Image  not null,
	WordName nvarchar(256) not null
);
go

use Georeports;
go


alter table DocumentsReports
	add constraint PK_SessionId primary key (Id),
		constraint FK_CatalogId foreign key (CatalogsId) references CatalogsReports(Id)
;


go

