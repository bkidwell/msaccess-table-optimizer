msaccess-table-optimizer
========================

Brendan Kidwell

17 December 2007

This module is a quick and dirty procedure for optimizing the sizes of Text fields in Access. One important use for this module would be cleaning up tables imported from external sources where Access defaults to very large field sizes.

The report script in this module creates a table called `_TableOptimizer` that lists the field size, size of shortest value, and size of longest value in the table you specify. Then if you want to resize any text fields, you can edit the report table and run the resizing script.

Usage
-----

1. Create a new module called `TableOptimizer`. Edit that module in the Visual Basic window (ALT-F11) and paste this code, replacing the stub code you find there initially.

2. Go to the Immediate window (CTRL-G) and run the method `TableOptimizer.MakeReport`. Enter the name of the table you want to analyze. The script will catalog the fields and their minimum and maximum sizes. WARNING: This will create a table called `_TableOptimizer`. If you already have such a table in your database, you must modify the code here appropriately.

3. Go to the Database window and open the `_TableOptimizer` window. Optionally filter and sort the table. (If you have run the script on more than one table, you will probably want to filter the `_TableOptimizer` table to show only records for that table.) Look at the `Size`, `Shortest_Value`, and `Longest_Value` numbers given for each Text field.

4. Enter a new value in the `New_Size` column for any Text fields you want to resize. Leave `New_Size` empty for any other fields.

5. Go to the Immediate window (CTRL-G) and run the method `TableOptimizer.ChangeSizes`. Again, enter the name of the table you're working on. The script will run a series of `ALTER TABLE` SQL statements to resize the Text fields as you specified. If there are any fields whose type you want to change (for example, Memo to Text), you must do this MANUALLY from the Design view of the target table.

