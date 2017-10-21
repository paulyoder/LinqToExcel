### 1.11.0

* Refactorings
  * Remove dependency on log4net.

### 1.10.1

* Refactorings
  * Manually removing log4net.dll from the package lib folder since it's not needed with the NuGet dependency on log4net

### 1.10

* Refactorings
  * Added Log4Net as a dependency in the NuGet file

### 1.9

* Enhancements
  * Added support for named ranges (by nkilian)

### 1.8.1

* Refactorings
  * Referencing Log4Net through its NuGet package

### 1.8

* enhancements
  * added **UsePersistentConnection** option to re-use the same connection for multiple queries. (by acorkery)
  * added **ReadOnly** option to open the file in readonly mode
