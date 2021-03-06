﻿CREATE DATABASE [ExcelFingerPrint]
GO

USE [ExcelFingerPrint]
GO
/****** Object:  Table [dbo].[FingerPrintData]    Script Date: 8/6/2020 4:16:42 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FingerPrintData](
	[Id] [nvarchar](255) NOT NULL,
	[GuestID] [nvarchar](50) NULL,
	[CardNo] [nvarchar](50) NULL,
	[GuestName] [nvarchar](255) NULL,
	[Department] [nvarchar](500) NULL,
	[Date] [nvarchar](50) NULL,
	[Time] [datetime] NULL,
	[EntryDoor] [nvarchar](50) NULL,
	[EventDescription] [nvarchar](50) NULL,
	[VerificationSource] [nvarchar](50) NULL,
PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
