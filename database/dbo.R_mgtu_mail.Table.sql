/****** Object:  Таблица исходящих писем в МГТУ ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[R_mgtu_mail](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[dt_z] [datetime] NULL CONSTRAINT [DF_R_mgtu_mail_dt_z]  DEFAULT (getdate()),
	[name_oi] [varchar](100) NULL CONSTRAINT [DF_R_mgtu_mail_name_oi]  DEFAULT ([dbo].[sp_get_namepol]()),
	[suser] [varchar](100) NULL CONSTRAINT [DF_R_mgtu_mail_suser]  DEFAULT ((suser_name()+' : ')+host_name()),
	[file_name] [varchar](500) NULL,
	[name_otch] [varchar](100) NULL,
	[file_guid] [varchar](100) NULL,
	[operator_name] [varchar](100) NULL,
	[operator_sign] [varchar](max) NULL,
	[processed] [datetime] NULL,
	[hash_file] [varchar](255) NULL,
	[message] [varchar](255) NULL,
	[path_file] [varchar](250) NULL,
	[dt_file] [datetime] NULL,
	[status] [nchar](3) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO


