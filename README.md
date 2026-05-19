# Movie Dashboard
## Table of Content
[Problem Statement](#problem-statement)  
[Data Source](#data-source)  
[Tools](#tools)  
[Data Cleaning and Preparation](#data-cleaning-and-preparation)  
[Dashboard](#dashboard)  
[M Code](#m-code)  
[Recomendations](#recomendations)  
### Problem Statement
Netflix wants to better understand which movies they should produce next, including the most suitable actors and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres. Your task is to build an Excel dashboard that provides insights into this dataset. The dashboard should help identify:
- The best-performing actors,
- The top movies based on box office metrics,
- Director performance,
- Genre trends,
- Seasonal patterns in movie performance,
- Any additional insights that can guide future production decisions.

The final dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.
### Data Source
Movie dataset: The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about each movie’s performance (box office and budget), actors, directors, etc.
You can download the original datasource here: [Movie Dataset Excel file](https://github.com/user-attachments/files/28031078/Movie_Data.xlsx)

### Tools
1. Power Query - for Data Cleaning and Data Transformation
2. Excel - for Data Analysis
3. Pivot Tables - for Creating the dashboard and Visualisations
   
### Data Cleaning and Preparation
- Data loading and inspection
- Handling errors and missing values
- Data cleaning and formatting. The excel file after the data cleaning and preparation process can be downloaded here: [Movie Dashboard](https://github.com/user-attachments/files/28029748/Movie_Dashboard.xlsx)

### Dashboard
<img width="1194" height="463" alt="image" src="https://github.com/user-attachments/assets/f65621aa-f831-4bf5-845e-99a0f210ad35" />

### M Code
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\yuriy\OneDrive\Рабочий стол\мама\PQ HW2\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres" = Table.ExpandTableColumn(#"Merged Queries", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Genres",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries1" = Table.NestedJoin(#"Reordered Columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres1" = Table.ExpandTableColumn(#"Merged Queries1", "Genres", {"Genre"}, {"Genres.Genre.1"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Expanded Genres1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres.Genre.1", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns1",{{"Genres.Genre.1", "Genres_Second"}, {"Genres.Genre", "Genres"}}),
    #"Merged Queries2" = Table.NestedJoin(#"Renamed Columns", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Expanded Directors" = Table.ExpandTableColumn(#"Merged Queries2", "Directors", {"Director"}, {"Directors.Director"}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Expanded Directors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres", "Genre_Second_ID", "Genres_Second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Reordered Columns2",{{"Directors.Director", "Directors"}}),
    #"Merged Queries3" = Table.NestedJoin(#"Renamed Columns1", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors" = Table.ExpandTableColumn(#"Merged Queries3", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Reordered Columns3" = Table.ReorderColumns(#"Expanded Actors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres", "Genre_Second_ID", "Genres_Second", "Director_First_ID", "Directors", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns3",{{"Actors.Actor", "Actors"}}),
    #"Merged Queries4" = Table.NestedJoin(#"Renamed Columns2", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors.1", JoinKind.LeftOuter),
    #"Expanded Actors.1" = Table.ExpandTableColumn(#"Merged Queries4", "Actors.1", {"Actor"}, {"Actors.1.Actor"}),
    #"Reordered Columns4" = Table.ReorderColumns(#"Expanded Actors.1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres", "Genre_Second_ID", "Genres_Second", "Director_First_ID", "Directors", "Cast_First_ID", "Actors", "Cast_Second_ID", "Actors.1.Actor", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Reordered Columns4",{{"Actors.1.Actor", "Actors_2"}}),
    #"Merged Queries5" = Table.NestedJoin(#"Renamed Columns3", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors.1", JoinKind.LeftOuter),
    #"Expanded Actors.2" = Table.ExpandTableColumn(#"Merged Queries5", "Actors.1", {"Actor"}, {"Actors.1.Actor"}),
    #"Reordered Columns5" = Table.ReorderColumns(#"Expanded Actors.2",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres", "Genre_Second_ID", "Genres_Second", "Director_First_ID", "Directors", "Cast_First_ID", "Actors", "Cast_Second_ID", "Actors_2", "Cast_Third_ID", "Actors.1.Actor", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Reordered Columns5",{{"Actors.1.Actor", "Actors_3"}}),
    #"Merged Queries6" = Table.NestedJoin(#"Renamed Columns4", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors.1", JoinKind.LeftOuter),
    #"Expanded Actors.3" = Table.ExpandTableColumn(#"Merged Queries6", "Actors.1", {"Actor"}, {"Actors.1.Actor"}),
    #"Reordered Columns6" = Table.ReorderColumns(#"Expanded Actors.3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres", "Genre_Second_ID", "Genres_Second", "Director_First_ID", "Directors", "Cast_First_ID", "Actors", "Cast_Second_ID", "Actors_2", "Cast_Third_ID", "Actors_3", "Cast_Fourth_ID", "Actors.1.Actor", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns5" = Table.RenameColumns(#"Reordered Columns6",{{"Actors.1.Actor", "Actors_4"}}),
    #"Merged Queries7" = Table.NestedJoin(#"Renamed Columns5", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors.1", JoinKind.LeftOuter),
    #"Expanded Actors.4" = Table.ExpandTableColumn(#"Merged Queries7", "Actors.1", {"Actor"}, {"Actors.1.Actor"}),
    #"Reordered Columns7" = Table.ReorderColumns(#"Expanded Actors.4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres", "Genre_Second_ID", "Genres_Second", "Director_First_ID", "Directors", "Cast_First_ID", "Actors", "Cast_Second_ID", "Actors_2", "Cast_Third_ID", "Actors_3", "Cast_Fourth_ID", "Actors_4", "Cast_Fifth_ID", "Actors.1.Actor", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns6" = Table.RenameColumns(#"Reordered Columns7",{{"Actors.1.Actor", "Actors_5"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns6",{{"Release Date", type date}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each true),
    #"Changed Type3" = Table.TransformColumnTypes(#"Filtered Rows",{{"ROI", Percentage.Type}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Changed Type3",{{"ROI", Percentage.Type}})
in
    #"Changed Type2"
```
### Recomendations
Based on box office revenue, Netflix should continue focusing on high-performing genres such as Action, Comedy, Drama, Sci-Fi, and Adventure.

<img width="532" height="164" alt="image" src="https://github.com/user-attachments/assets/5da050a1-3c54-4ea3-b75a-337ecf8e4173" />

However, Horror movies demonstrated the highest ROI, suggesting they can be a highly profitable option with lower production budgets.

<img width="643" height="165" alt="image" src="https://github.com/user-attachments/assets/b38ed67b-8586-463a-abfe-e23538bd4c91" />



