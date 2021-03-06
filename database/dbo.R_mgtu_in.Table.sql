/****** Object:  Таблица входящих писем от МГТУ ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[R_mgtu_in](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[file_name] [varchar](255) NULL,
	[file_path] [varchar](255) NULL,
	[received_from] [varchar](100) NULL,
	[subject] [varchar](255) NULL,
	[file_id] [varchar](100) NULL,
	[sender] [varchar](100) NULL,
	[received_date] [datetime] NULL,
	[processed_date] [datetime] NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO


