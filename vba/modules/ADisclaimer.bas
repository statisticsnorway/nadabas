Attribute VB_Name = "ADisclaimer"
'
' NADABAS is developed by Søren Netterstrøm, Statistics Denmark
' in cooperation with Jan Redeby, Statistics Sweden
' for the Scandinavian Cooperation project with INE, Mozambique.
'
' It may be freely used by any official statistical office
' and adapted in any way you feel
' The authors, Statistics Denmark and Statics Sweden take no responsibility
' for this product.

' NADABAS is now developed by Jarle Kvile, Statistics Norway.

' If you have comments or suggestions, please contact the author
' jkv@ssb.no
'
' Please keep reference to the original source if you modify the system
'
' The first version of NADABAS was produced in April 2004.
'

'**********************************************************************************************
' Module Name: [ModuleName]
' Description: [Provide a brief description of the module's purpose and functionality.
'               For example: This module handles the initialization and management of
'               language settings in NADABAS.]
'
' Author: [Your Name]
' Created: [Creation Date]
' Last Updated: [Last Update Date]
'
' Dependencies:
'   - [List any dependencies, such as other modules, external libraries, or specific settings.]
'
' Key Procedures:
'   - [ProcedureName1]: [Brief description of what it does.]
'   - [ProcedureName2]: [Brief description of what it does.]
'
' Notes:
'   - [Optional: Include any additional notes or assumptions.]
'
'**********************************************************************************************


'**********************************************************************************************
' Class Name: clsDB
' Description: Represents a database connection in NADABAS, enabling operations such as querying
'              and updating tables. Includes methods for opening, closing, and executing commands.
'
' Author: Jane Doe
' Created: February 1, 2024
' Last Updated: March 5, 2024
'
' Properties:
'   - ConnectionString: Stores the database connection string.
'   - IsConnected: Returns whether the database is currently connected.
'
' Methods:
'   - OpenConnection: Establishes a connection to the database.
'   - CloseConnection: Closes the current database connection.
'   - ExecuteQuery: Executes a SQL query and returns the results.
'
' Events:
'   - OnConnectionError: Triggered when a connection error occurs.
'
' Notes:
'   - Ensure that all connections are closed properly to avoid resource leaks.
'
'**********************************************************************************************
