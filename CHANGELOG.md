# Changelog

### 2.0.0 (PRERELEASE-2)
* Allow connstring suppression of TransactionScope. Thanks **@freakingawesome**.
* Allow lazy queries using yield statements. Thanks **@freakingawesome**.
* Add means to gather exceptions instead of throwing. Thanks **@freakingawesome**.
* Implemented Source Link for easier debugging. Thanks **@MagicAndre1981**.

### 2.0.0 (PRERELEASE)

* Remove support for Microsoft Jet as it was deprecated many years ago, and only works in 32-bit applications.
  * MS Ace driver is now a *requirement*.
* Target AnyCPU
* Target .Net Framework v3.5, v4.5.1, and v4.6.
* Fix `ExcelQueryFactory` not being disposed properly.
* Fix incorrect worksheet names that contain a `$`.
* Remotion.Linq updated, and no longer bundled with project.
* Added support for unary expressions in Linq aggregate functions.
* Added support for primitive value type results to be cast to `Nullable<T>` counterparts.
* Fix "item already inserted" issue in AddTransformation. Thanks **@tkestowicz**.
* Throw Exception with row number and column name/number. Thanks **@achvaicer**.
* Added a method of gathering "unmapped cells". Thanks **@freakingawesome**.

* Other notes:
  * Thanks to **@cuongtranba** for helping to move the project to Nunit.


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
