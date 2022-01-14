DROP TABLE arreas;
CREATE TABLE [dbo].[arrears] (
    [name_groups]     VARCHAR (10)  NOT NULL,
    [name_students]   VARCHAR (70)  NOT NULL,
    [numberOfSemestr] VARCHAR (3)   NOT NULL,
    [semiannual]          VARCHAR (10) NOT NULL,
    [lesson]      VARCHAR (200)  NOT NULL,
    [departament]     VARCHAR (10)  NOT NULL,
    [typeOfControl]   VARCHAR (10)  NOT NULL
);