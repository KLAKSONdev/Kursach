USE [master]
GO
/****** Object:  Database [vsstu]    Script Date: 10.03.2026 23:23:40 ******/
CREATE DATABASE [vsstu]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'vsstu', FILENAME = N'C:\Users\Клаксон\vsstu.mdf' , SIZE = 73728KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'vsstu_log', FILENAME = N'C:\Users\Клаксон\vsstu_log.ldf' , SIZE = 73728KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT, LEDGER = OFF
GO
ALTER DATABASE [vsstu] SET COMPATIBILITY_LEVEL = 170
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [vsstu].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [vsstu] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [vsstu] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [vsstu] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [vsstu] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [vsstu] SET ARITHABORT OFF 
GO
ALTER DATABASE [vsstu] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [vsstu] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [vsstu] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [vsstu] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [vsstu] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [vsstu] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [vsstu] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [vsstu] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [vsstu] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [vsstu] SET  DISABLE_BROKER 
GO
ALTER DATABASE [vsstu] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [vsstu] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [vsstu] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [vsstu] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [vsstu] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [vsstu] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [vsstu] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [vsstu] SET RECOVERY SIMPLE 
GO
ALTER DATABASE [vsstu] SET  MULTI_USER 
GO
ALTER DATABASE [vsstu] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [vsstu] SET DB_CHAINING OFF 
GO
ALTER DATABASE [vsstu] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [vsstu] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [vsstu] SET DELAYED_DURABILITY = DISABLED 
GO
ALTER DATABASE [vsstu] SET ACCELERATED_DATABASE_RECOVERY = OFF  
GO
ALTER DATABASE [vsstu] SET OPTIMIZED_LOCKING = OFF 
GO
ALTER DATABASE [vsstu] SET QUERY_STORE = ON
GO
ALTER DATABASE [vsstu] SET QUERY_STORE (OPERATION_MODE = READ_WRITE, CLEANUP_POLICY = (STALE_QUERY_THRESHOLD_DAYS = 30), DATA_FLUSH_INTERVAL_SECONDS = 900, INTERVAL_LENGTH_MINUTES = 60, MAX_STORAGE_SIZE_MB = 1000, QUERY_CAPTURE_MODE = AUTO, SIZE_BASED_CLEANUP_MODE = AUTO, MAX_PLANS_PER_QUERY = 200, WAIT_STATS_CAPTURE_MODE = ON)
GO
USE [vsstu]
GO
/****** Object:  Table [dbo].[Absences]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Absences](
	[AbsenceID] [int] IDENTITY(1,1) NOT NULL,
	[StudentID] [int] NULL,
	[StartDate] [date] NOT NULL,
	[EndDate] [date] NULL,
	[Days] [int] NULL,
	[Reason] [nvarchar](max) NULL,
	[IsRespectful] [bit] NULL,
	[DocumentProof] [nvarchar](255) NULL,
	[Notes] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[AbsenceID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AcademicPerformance]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AcademicPerformance](
	[PerformanceID] [int] IDENTITY(1,1) NOT NULL,
	[StudentID] [int] NULL,
	[Semester] [int] NULL,
	[SubjectName] [nvarchar](200) NOT NULL,
	[Grade] [int] NULL,
	[GradeType] [nvarchar](50) NULL,
	[Date] [date] NULL,
	[Teacher] [nvarchar](200) NULL,
	[Notes] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[PerformanceID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Achievements]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Achievements](
	[AchievementID] [int] IDENTITY(1,1) NOT NULL,
	[StudentID] [int] NULL,
	[AchievementTypeID] [int] NULL,
	[AchievementName] [nvarchar](200) NOT NULL,
	[Date] [date] NULL,
	[Level] [nvarchar](50) NULL,
	[Place] [nvarchar](50) NULL,
	[Document] [nvarchar](255) NULL,
	[Notes] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[AchievementID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AchievementTypes]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AchievementTypes](
	[AchievementTypeID] [int] IDENTITY(1,1) NOT NULL,
	[TypeName] [nvarchar](100) NOT NULL,
	[Description] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[AchievementTypeID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Attendance]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Attendance](
	[AttendanceID] [int] IDENTITY(1,1) NOT NULL,
	[StudentID] [int] NULL,
	[Date] [date] NOT NULL,
	[Subject] [nvarchar](200) NULL,
	[IsPresent] [bit] NULL,
	[IsLate] [bit] NULL,
	[IsAbsentRespectful] [bit] NULL,
	[IsAbsentDisrespectful] [bit] NULL,
	[Hours] [int] NULL,
	[Reason] [nvarchar](max) NULL,
	[MarkedBy] [int] NULL,
PRIMARY KEY CLUSTERED 
(
	[AttendanceID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CuratorGroups]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CuratorGroups](
	[CuratorGroupID] [int] IDENTITY(1,1) NOT NULL,
	[CuratorID] [int] NULL,
	[GroupID] [int] NULL,
	[AssignmentDate] [date] NULL,
	[EndDate] [date] NULL,
	[IsActive] [bit] NULL,
PRIMARY KEY CLUSTERED 
(
	[CuratorGroupID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CuratorHourAttendance]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CuratorHourAttendance](
	[AttendanceID] [int] IDENTITY(1,1) NOT NULL,
	[HourID] [int] NULL,
	[StudentID] [int] NULL,
	[IsPresent] [bit] NULL,
	[Notes] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[AttendanceID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CuratorHours]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CuratorHours](
	[HourID] [int] IDENTITY(1,1) NOT NULL,
	[CuratorID] [int] NULL,
	[GroupID] [int] NULL,
	[Date] [date] NOT NULL,
	[Topic] [nvarchar](500) NULL,
	[PresentCount] [int] NULL,
	[Notes] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[HourID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Curators]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Curators](
	[CuratorID] [int] IDENTITY(1,1) NOT NULL,
	[LastName] [nvarchar](50) NOT NULL,
	[FirstName] [nvarchar](50) NOT NULL,
	[MiddleName] [nvarchar](50) NULL,
	[BirthDate] [date] NULL,
	[Phone] [nvarchar](20) NULL,
	[Email] [nvarchar](100) NULL,
	[Position] [nvarchar](100) NULL,
	[Department] [nvarchar](100) NULL,
	[PhotoPath] [nvarchar](255) NULL,
	[HireDate] [date] NULL,
	[Login] [nvarchar](50) NULL,
	[PasswordHash] [nvarchar](255) NULL,
	[IsActive] [bit] NULL,
	[CreatedAt] [datetime2](7) NULL,
PRIMARY KEY CLUSTERED 
(
	[CuratorID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[Login] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DisciplinaryRecords]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DisciplinaryRecords](
	[RecordID] [int] IDENTITY(1,1) NOT NULL,
	[StudentID] [int] NULL,
	[RecordType] [nvarchar](20) NULL,
	[Date] [date] NOT NULL,
	[Reason] [nvarchar](max) NULL,
	[IssuedBy] [nvarchar](200) NULL,
	[Notes] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[RecordID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[EventParticipation]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EventParticipation](
	[ParticipationID] [int] IDENTITY(1,1) NOT NULL,
	[EventID] [int] NULL,
	[StudentID] [int] NULL,
	[ParticipationStatus] [nvarchar](50) NULL,
	[Role] [nvarchar](100) NULL,
	[Result] [nvarchar](max) NULL,
	[Notes] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[ParticipationID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Events]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Events](
	[EventID] [int] IDENTITY(1,1) NOT NULL,
	[EventName] [nvarchar](200) NOT NULL,
	[EventTypeID] [int] NULL,
	[EventDate] [date] NULL,
	[EventTime] [time](7) NULL,
	[Location] [nvarchar](200) NULL,
	[Organizer] [nvarchar](200) NULL,
	[Description] [nvarchar](max) NULL,
	[IsRequired] [bit] NULL,
	[CreatedAt] [datetime2](7) NULL,
PRIMARY KEY CLUSTERED 
(
	[EventID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[EventTypes]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EventTypes](
	[EventTypeID] [int] IDENTITY(1,1) NOT NULL,
	[EventTypeName] [nvarchar](100) NOT NULL,
	[Description] [nvarchar](max) NULL,
	[Category] [nvarchar](50) NULL,
PRIMARY KEY CLUSTERED 
(
	[EventTypeID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Faculties]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Faculties](
	[FacultyID] [int] IDENTITY(1,1) NOT NULL,
	[FacultyName] [nvarchar](100) NOT NULL,
	[FacultyCode] [nvarchar](10) NULL,
	[Description] [nvarchar](max) NULL,
	[CreatedAt] [datetime2](7) NULL,
PRIMARY KEY CLUSTERED 
(
	[FacultyID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[FacultyCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FamilyTypes]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FamilyTypes](
	[FamilyTypeID] [int] IDENTITY(1,1) NOT NULL,
	[FamilyTypeName] [nvarchar](50) NOT NULL,
	[Description] [nvarchar](max) NULL,
	[IsActive] [bit] NULL,
PRIMARY KEY CLUSTERED 
(
	[FamilyTypeID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Git]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Git](
	[Git] [nchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Groups]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Groups](
	[GroupID] [int] IDENTITY(1,1) NOT NULL,
	[GroupName] [nvarchar](20) NOT NULL,
	[SpecialtyID] [int] NULL,
	[Course] [int] NULL,
	[AcademicYear] [nvarchar](9) NULL,
	[StudentCount] [int] NULL,
	[FormOfEducation] [nvarchar](20) NULL,
	[Language] [nvarchar](20) NULL,
	[CreatedAt] [datetime2](7) NULL,
PRIMARY KEY CLUSTERED 
(
	[GroupID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[GroupName] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[HealthStatus]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[HealthStatus](
	[HealthStatusID] [int] IDENTITY(1,1) NOT NULL,
	[HealthStatusName] [nvarchar](50) NOT NULL,
	[Description] [nvarchar](max) NULL,
	[IsActive] [bit] NULL,
PRIMARY KEY CLUSTERED 
(
	[HealthStatusID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[NegativeTraits]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NegativeTraits](
	[NegativeTraitID] [int] IDENTITY(1,1) NOT NULL,
	[TraitName] [nvarchar](100) NOT NULL,
	[Description] [nvarchar](max) NULL,
	[Category] [nvarchar](50) NULL,
PRIMARY KEY CLUSTERED 
(
	[NegativeTraitID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Parents]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Parents](
	[ParentID] [int] IDENTITY(1,1) NOT NULL,
	[LastName] [nvarchar](50) NOT NULL,
	[FirstName] [nvarchar](50) NOT NULL,
	[MiddleName] [nvarchar](50) NULL,
	[Relationship] [nvarchar](50) NULL,
	[Phone] [nvarchar](20) NULL,
	[Email] [nvarchar](100) NULL,
	[WorkPlace] [nvarchar](200) NULL,
	[WorkPosition] [nvarchar](100) NULL,
	[Address] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[ParentID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Passwords]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Passwords](
	[PasswordID] [int] IDENTITY(1,1) NOT NULL,
	[UserType] [nvarchar](20) NOT NULL,
	[UserID] [int] NOT NULL,
	[PasswordHash] [nvarchar](255) NOT NULL,
	[CreatedAt] [datetime2](7) NULL,
	[UpdatedAt] [datetime2](7) NULL,
PRIMARY KEY CLUSTERED 
(
	[PasswordID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PositiveTraits]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PositiveTraits](
	[PositiveTraitID] [int] IDENTITY(1,1) NOT NULL,
	[TraitName] [nvarchar](100) NOT NULL,
	[Description] [nvarchar](max) NULL,
	[Category] [nvarchar](50) NULL,
PRIMARY KEY CLUSTERED 
(
	[PositiveTraitID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Schedule]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Schedule](
	[ScheduleID] [int] IDENTITY(1,1) NOT NULL,
	[GroupID] [int] NOT NULL,
	[DayOfWeek] [int] NOT NULL,
	[Subject] [nvarchar](200) NOT NULL,
	[Teacher] [nvarchar](200) NULL,
	[StartTime] [time](7) NULL,
	[EndTime] [time](7) NULL,
	[IsActive] [bit] NULL,
	[CreatedAt] [datetime2](7) NULL,
	[UpdatedAt] [datetime2](7) NULL,
PRIMARY KEY CLUSTERED 
(
	[ScheduleID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Specialties]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Specialties](
	[SpecialtyID] [int] IDENTITY(1,1) NOT NULL,
	[SpecialtyCode] [nvarchar](20) NOT NULL,
	[SpecialtyName] [nvarchar](200) NOT NULL,
	[FacultyID] [int] NULL,
	[EducationLevel] [nvarchar](50) NULL,
	[DurationYears] [int] NULL,
	[Description] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[SpecialtyID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[SpecialtyCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[StudentCharacteristics]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[StudentCharacteristics](
	[CharacteristicID] [int] IDENTITY(1,1) NOT NULL,
	[StudentID] [int] NOT NULL,
	[CharacteristicText] [nvarchar](max) NULL,
	[CreatedDate] [datetime] NULL,
	[UpdatedDate] [datetime] NULL,
	[CreatedBy] [int] NULL,
 CONSTRAINT [PK_StudentCharacteristics] PRIMARY KEY CLUSTERED 
(
	[CharacteristicID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[StudentParents]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[StudentParents](
	[StudentParentID] [int] IDENTITY(1,1) NOT NULL,
	[StudentID] [int] NULL,
	[ParentID] [int] NULL,
	[IsPrimaryContact] [bit] NULL,
PRIMARY KEY CLUSTERED 
(
	[StudentParentID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Students]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Students](
	[StudentID] [int] IDENTITY(1,1) NOT NULL,
	[LastName] [nvarchar](50) NOT NULL,
	[FirstName] [nvarchar](50) NOT NULL,
	[MiddleName] [nvarchar](50) NULL,
	[GroupID] [int] NULL,
	[StudentCardNumber] [nvarchar](20) NULL,
	[PersonalNumber] [nvarchar](20) NULL,
	[BirthDate] [date] NULL,
	[BirthPlace] [nvarchar](200) NULL,
	[Gender] [nvarchar](1) NULL,
	[Nationality] [nvarchar](50) NULL,
	[Citizenship] [nvarchar](50) NULL,
	[EducationBefore] [nvarchar](100) NULL,
	[EducationDocument] [nvarchar](100) NULL,
	[PhotoPath] [nvarchar](255) NULL,
	[Phone] [nvarchar](20) NULL,
	[Email] [nvarchar](100) NULL,
	[ParentsPhone] [nvarchar](20) NULL,
	[RegistrationAddress] [nvarchar](max) NULL,
	[ResidentialAddress] [nvarchar](max) NULL,
	[HealthStatusID] [int] NULL,
	[FamilyTypeID] [int] NULL,
	[IsOrphan] [bit] NULL,
	[IsDisabled] [bit] NULL,
	[IsFromLargeFamily] [bit] NULL,
	[IsLowIncome] [bit] NULL,
	[IsEmployed] [bit] NULL,
	[WorkPlace] [nvarchar](200) NULL,
	[WorkPosition] [nvarchar](100) NULL,
	[Login] [nvarchar](50) NULL,
	[PasswordHash] [nvarchar](255) NULL,
	[IsActive] [bit] NULL,
	[EnrollmentDate] [date] NULL,
	[GraduationDate] [date] NULL,
	[CreatedAt] [datetime2](7) NULL,
	[UpdatedAt] [datetime2](7) NULL,
	[Photo] [varbinary](max) NULL,
	[IsHeadman] [bit] NULL,
 CONSTRAINT [PK__Students__32C52A79A7B7651B] PRIMARY KEY CLUSTERED 
(
	[StudentID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[StudentTraits]    Script Date: 10.03.2026 23:23:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[StudentTraits](
	[TraitID] [int] IDENTITY(1,1) NOT NULL,
	[StudentID] [int] NULL,
	[PositiveTraitID] [int] NULL,
	[NegativeTraitID] [int] NULL,
	[Notes] [nvarchar](max) NULL,
	[MarkedBy] [int] NULL,
	[Date] [date] NULL,
PRIMARY KEY CLUSTERED 
(
	[TraitID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_Passwords_User]    Script Date: 10.03.2026 23:23:40 ******/
CREATE NONCLUSTERED INDEX [IX_Passwords_User] ON [dbo].[Passwords]
(
	[UserType] ASC,
	[UserID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [IX_StudentCharacteristics_StudentID]    Script Date: 10.03.2026 23:23:41 ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_StudentCharacteristics_StudentID] ON [dbo].[StudentCharacteristics]
(
	[StudentID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_Students_Login]    Script Date: 10.03.2026 23:23:41 ******/
CREATE NONCLUSTERED INDEX [IX_Students_Login] ON [dbo].[Students]
(
	[Login] ASC
)
WHERE ([Login] IS NOT NULL)
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_Students_PersonalNumber]    Script Date: 10.03.2026 23:23:41 ******/
CREATE NONCLUSTERED INDEX [IX_Students_PersonalNumber] ON [dbo].[Students]
(
	[PersonalNumber] ASC
)
WHERE ([PersonalNumber] IS NOT NULL)
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_Students_StudentCardNumber]    Script Date: 10.03.2026 23:23:41 ******/
CREATE NONCLUSTERED INDEX [IX_Students_StudentCardNumber] ON [dbo].[Students]
(
	[StudentCardNumber] ASC
)
WHERE ([StudentCardNumber] IS NOT NULL)
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Absences] ADD  DEFAULT ((1)) FOR [IsRespectful]
GO
ALTER TABLE [dbo].[Attendance] ADD  DEFAULT ((1)) FOR [IsPresent]
GO
ALTER TABLE [dbo].[Attendance] ADD  DEFAULT ((0)) FOR [IsLate]
GO
ALTER TABLE [dbo].[Attendance] ADD  DEFAULT ((0)) FOR [IsAbsentRespectful]
GO
ALTER TABLE [dbo].[Attendance] ADD  DEFAULT ((0)) FOR [IsAbsentDisrespectful]
GO
ALTER TABLE [dbo].[Attendance] ADD  DEFAULT ((2)) FOR [Hours]
GO
ALTER TABLE [dbo].[CuratorGroups] ADD  DEFAULT (getdate()) FOR [AssignmentDate]
GO
ALTER TABLE [dbo].[CuratorGroups] ADD  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[CuratorHourAttendance] ADD  DEFAULT ((1)) FOR [IsPresent]
GO
ALTER TABLE [dbo].[Curators] ADD  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[Curators] ADD  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[Events] ADD  DEFAULT ((0)) FOR [IsRequired]
GO
ALTER TABLE [dbo].[Events] ADD  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[Faculties] ADD  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[FamilyTypes] ADD  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[Groups] ADD  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[HealthStatus] ADD  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[Passwords] ADD  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[Passwords] ADD  DEFAULT (getdate()) FOR [UpdatedAt]
GO
ALTER TABLE [dbo].[Schedule] ADD  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[Schedule] ADD  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[Schedule] ADD  DEFAULT (getdate()) FOR [UpdatedAt]
GO
ALTER TABLE [dbo].[StudentParents] ADD  DEFAULT ((0)) FOR [IsPrimaryContact]
GO
ALTER TABLE [dbo].[Students] ADD  CONSTRAINT [DF__Students__IsOrph__73BA3083]  DEFAULT ((0)) FOR [IsOrphan]
GO
ALTER TABLE [dbo].[Students] ADD  CONSTRAINT [DF__Students__IsDisa__74AE54BC]  DEFAULT ((0)) FOR [IsDisabled]
GO
ALTER TABLE [dbo].[Students] ADD  CONSTRAINT [DF__Students__IsFrom__75A278F5]  DEFAULT ((0)) FOR [IsFromLargeFamily]
GO
ALTER TABLE [dbo].[Students] ADD  CONSTRAINT [DF__Students__IsLowI__76969D2E]  DEFAULT ((0)) FOR [IsLowIncome]
GO
ALTER TABLE [dbo].[Students] ADD  CONSTRAINT [DF__Students__IsEmpl__778AC167]  DEFAULT ((0)) FOR [IsEmployed]
GO
ALTER TABLE [dbo].[Students] ADD  CONSTRAINT [DF__Students__IsActi__787EE5A0]  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[Students] ADD  CONSTRAINT [DF__Students__Create__797309D9]  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[Students] ADD  CONSTRAINT [DF__Students__Update__7A672E12]  DEFAULT (getdate()) FOR [UpdatedAt]
GO
ALTER TABLE [dbo].[StudentTraits] ADD  DEFAULT (getdate()) FOR [Date]
GO
ALTER TABLE [dbo].[Absences]  WITH CHECK ADD  CONSTRAINT [FK__Absences__Studen__2BFE89A6] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[Absences] CHECK CONSTRAINT [FK__Absences__Studen__2BFE89A6]
GO
ALTER TABLE [dbo].[AcademicPerformance]  WITH CHECK ADD  CONSTRAINT [FK__AcademicP__Stude__08B54D69] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[AcademicPerformance] CHECK CONSTRAINT [FK__AcademicP__Stude__08B54D69]
GO
ALTER TABLE [dbo].[Achievements]  WITH CHECK ADD FOREIGN KEY([AchievementTypeID])
REFERENCES [dbo].[AchievementTypes] ([AchievementTypeID])
GO
ALTER TABLE [dbo].[Achievements]  WITH CHECK ADD  CONSTRAINT [FK__Achieveme__Stude__1CBC4616] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[Achievements] CHECK CONSTRAINT [FK__Achieveme__Stude__1CBC4616]
GO
ALTER TABLE [dbo].[Attendance]  WITH CHECK ADD FOREIGN KEY([MarkedBy])
REFERENCES [dbo].[Curators] ([CuratorID])
GO
ALTER TABLE [dbo].[Attendance]  WITH CHECK ADD  CONSTRAINT [FK__Attendanc__Stude__10566F31] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[Attendance] CHECK CONSTRAINT [FK__Attendanc__Stude__10566F31]
GO
ALTER TABLE [dbo].[CuratorGroups]  WITH CHECK ADD FOREIGN KEY([CuratorID])
REFERENCES [dbo].[Curators] ([CuratorID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[CuratorGroups]  WITH CHECK ADD FOREIGN KEY([GroupID])
REFERENCES [dbo].[Groups] ([GroupID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[CuratorHourAttendance]  WITH CHECK ADD FOREIGN KEY([HourID])
REFERENCES [dbo].[CuratorHours] ([HourID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[CuratorHourAttendance]  WITH CHECK ADD  CONSTRAINT [FK__CuratorHo__Stude__3493CFA7] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[CuratorHourAttendance] CHECK CONSTRAINT [FK__CuratorHo__Stude__3493CFA7]
GO
ALTER TABLE [dbo].[CuratorHours]  WITH CHECK ADD FOREIGN KEY([CuratorID])
REFERENCES [dbo].[Curators] ([CuratorID])
GO
ALTER TABLE [dbo].[CuratorHours]  WITH CHECK ADD FOREIGN KEY([GroupID])
REFERENCES [dbo].[Groups] ([GroupID])
GO
ALTER TABLE [dbo].[DisciplinaryRecords]  WITH CHECK ADD  CONSTRAINT [FK__Disciplin__Stude__282DF8C2] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[DisciplinaryRecords] CHECK CONSTRAINT [FK__Disciplin__Stude__282DF8C2]
GO
ALTER TABLE [dbo].[EventParticipation]  WITH CHECK ADD FOREIGN KEY([EventID])
REFERENCES [dbo].[Events] ([EventID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[EventParticipation]  WITH CHECK ADD  CONSTRAINT [FK__EventPart__Stude__19DFD96B] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[EventParticipation] CHECK CONSTRAINT [FK__EventPart__Stude__19DFD96B]
GO
ALTER TABLE [dbo].[Events]  WITH CHECK ADD FOREIGN KEY([EventTypeID])
REFERENCES [dbo].[EventTypes] ([EventTypeID])
GO
ALTER TABLE [dbo].[Groups]  WITH CHECK ADD FOREIGN KEY([SpecialtyID])
REFERENCES [dbo].[Specialties] ([SpecialtyID])
GO
ALTER TABLE [dbo].[Schedule]  WITH CHECK ADD FOREIGN KEY([GroupID])
REFERENCES [dbo].[Groups] ([GroupID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[Specialties]  WITH CHECK ADD FOREIGN KEY([FacultyID])
REFERENCES [dbo].[Faculties] ([FacultyID])
GO
ALTER TABLE [dbo].[StudentCharacteristics]  WITH CHECK ADD  CONSTRAINT [FK_StudentCharacteristics_Students] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[StudentCharacteristics] CHECK CONSTRAINT [FK_StudentCharacteristics_Students]
GO
ALTER TABLE [dbo].[StudentParents]  WITH CHECK ADD FOREIGN KEY([ParentID])
REFERENCES [dbo].[Parents] ([ParentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[StudentParents]  WITH CHECK ADD  CONSTRAINT [FK__StudentPa__Stude__02FC7413] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[StudentParents] CHECK CONSTRAINT [FK__StudentPa__Stude__02FC7413]
GO
ALTER TABLE [dbo].[Students]  WITH CHECK ADD  CONSTRAINT [FK__Students__Family__7D439ABD] FOREIGN KEY([FamilyTypeID])
REFERENCES [dbo].[FamilyTypes] ([FamilyTypeID])
GO
ALTER TABLE [dbo].[Students] CHECK CONSTRAINT [FK__Students__Family__7D439ABD]
GO
ALTER TABLE [dbo].[Students]  WITH CHECK ADD  CONSTRAINT [FK__Students__GroupI__7B5B524B] FOREIGN KEY([GroupID])
REFERENCES [dbo].[Groups] ([GroupID])
GO
ALTER TABLE [dbo].[Students] CHECK CONSTRAINT [FK__Students__GroupI__7B5B524B]
GO
ALTER TABLE [dbo].[Students]  WITH CHECK ADD  CONSTRAINT [FK__Students__Health__7C4F7684] FOREIGN KEY([HealthStatusID])
REFERENCES [dbo].[HealthStatus] ([HealthStatusID])
GO
ALTER TABLE [dbo].[Students] CHECK CONSTRAINT [FK__Students__Health__7C4F7684]
GO
ALTER TABLE [dbo].[StudentTraits]  WITH CHECK ADD FOREIGN KEY([MarkedBy])
REFERENCES [dbo].[Curators] ([CuratorID])
GO
ALTER TABLE [dbo].[StudentTraits]  WITH CHECK ADD FOREIGN KEY([NegativeTraitID])
REFERENCES [dbo].[NegativeTraits] ([NegativeTraitID])
GO
ALTER TABLE [dbo].[StudentTraits]  WITH CHECK ADD FOREIGN KEY([PositiveTraitID])
REFERENCES [dbo].[PositiveTraits] ([PositiveTraitID])
GO
ALTER TABLE [dbo].[StudentTraits]  WITH CHECK ADD  CONSTRAINT [FK__StudentTr__Stude__2180FB33] FOREIGN KEY([StudentID])
REFERENCES [dbo].[Students] ([StudentID])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[StudentTraits] CHECK CONSTRAINT [FK__StudentTr__Stude__2180FB33]
GO
ALTER TABLE [dbo].[AcademicPerformance]  WITH CHECK ADD CHECK  (([Grade]>=(1) AND [Grade]<=(5)))
GO
ALTER TABLE [dbo].[AcademicPerformance]  WITH CHECK ADD CHECK  (([Semester]>=(1) AND [Semester]<=(12)))
GO
ALTER TABLE [dbo].[DisciplinaryRecords]  WITH CHECK ADD CHECK  (([RecordType]='???????' OR [RecordType]='?????????' OR [RecordType]='?????????'))
GO
ALTER TABLE [dbo].[Groups]  WITH CHECK ADD CHECK  (([Course]>=(1) AND [Course]<=(6)))
GO
USE [master]
GO
ALTER DATABASE [vsstu] SET  READ_WRITE 
GO
