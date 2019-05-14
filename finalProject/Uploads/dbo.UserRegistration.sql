CREATE TABLE [dbo].[UserRegistration]
(
	[StaffCode]       VARCHAR (50)  NOT NULL,
    [StaffName]       VARCHAR (150) NOT NULL,
    [EmailId]         VARCHAR (50)  NOT NULL,
    [Designation]     VARCHAR (100) NOT NULL,
    [ContactNumber]   VARCHAR (50)  NOT NULL,
    [Department]      VARCHAR (50)  NOT NULL,
    [Password]        VARCHAR (50)  NOT NULL,
    [ConfirmPassword] VARCHAR (50)  NOT NULL,
    [Role]            VARCHAR (50)  NOT NULL,
    CONSTRAINT [PK_UserRegistration] PRIMARY KEY CLUSTERED ([StaffCode] ASC)
)
