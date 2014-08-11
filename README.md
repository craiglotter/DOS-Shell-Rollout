DOS Shell Rollout
=================

The problem in managing a large number of PCs is keeping them all uptodate and uniform in terms of the software provided on them. In most instances this is managed through a process of imaging, using various existing software solutions. However, rolling out images is particularly tedious, especially across large networks and for small inconsequential changes. DOS Shell Rollout provides a simple solution to providing for temporary updates to all networked PCs without having to rollout images.

The system works by utilising a database on the one end that is used to contain a number of DOS commands. A network PC on startup will log into the database, check which DOS statements have already been executed and execute those statements not yet implemented. All transactions are logged so that the system keeps itself up to date with what has and has not been applied.

The uses for this system is virtually limitless as almost anything can be done via DOS commands. For instance the quiet install of a new application can be done by executing the file from a open network server using the /passive switch or the simple removing of offending file or file hierarchies via targeted deltree or del commands.

Created by Craig Lotter, November 2005

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2003
Implements concepts such as threading, SQL programming and Database manipulation, Shell Scripting.
Level of Complexity: simple
