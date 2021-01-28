Accompanying code for https://rubberduckvba.wordpress.com/2020/04/22/secure-adodb/

---

# SecureADODB

## Purpose

A thin wrapper around the `ADODB` library, automating the creating of `ADODB.Connection`, `ADODB.Command`, `ADODB.Parameter` objects to securely communicate with a database using VBA. The `IDbConnection` is a *leaky abstraction* on purpose, to support more advanced scenarios and retain the full functionality of ADODB, while providing a simple-to-use API for the most common scenarios.

## References

This code uses early-bound reference to the `ADODB` library: the project where this code is imported must have a reference *Microsoft ActiveX Data Objects*. Alternatively, the ADODB-specific code could be modified to use late binding instead (pull requests welcome!), and then the project could be compiled without adding such a reference.

[**Rubberduck**](https://github.com/rubberduck-vba/Rubberduck) is not required to view the code, but using the *Code Explorer* toolwindow will make the structure much more apparent and easier to navigate:

![SecureADODB-ModuleView (Code Explorer)](https://user-images.githubusercontent.com/5751684/79779711-cf5a4c80-8308-11ea-90fd-2dd2a831b6c0.PNG)

Rubberduck is required for running the unit tests, but the modules under the `Test` folder (standard modules with a `@TestModule` annotation comment, and the `StubXxxxx` class modules too) are not needed for the code to be usable. Recommendation is to include the unit tests if you're planning to modify the code in any way.

## Usage

### UnitOfWork

The top-level API object is `UnitOfWork` and the `IUnitOfWork` interface. Given a `connString` connection string, and a `sql` command string, the simplest way to use it is to start with the `UnitOfWork.FromConnectionString` factory method - using a `With` block, we don't even need to declare the `IUnitOfWork` object:

```vb
With UnitOfWork.FromConnectionString(connString)
    'connection is open, a transaction is initiated.
    
    'IDbCommand.Execute returns a disconnected ADODB.Recordset:
    Dim results As ADODB.Recordset
    'simply use '?' ordinal parameters in the command string, and then provide a value for each '?' in the SQL:
    Set results = .Command.Execute("SELECT * FROM Table1 WHERE Field1 = ? AND Field2 = ?", 42, "Test")
        
    'if we don't need a recordset, we can use ExecuteNonQuery too:
    .Command.ExecuteNonQuery "INSERT INTO Table1 (DateInserted, Field1, Field2) VALUES (GETDATE(), ?, ?)", 42, "TEST"
        
    'if we only want to select a single value, we can use GetSingleValue:
    Dim result As Long
    result = .Command.GetSingleValue("SELECT Field2 FROM Table1 WHERE Field1 = ?", 42)
        
    'we are in a transaction, so we need to commit the changes - lest we lose them:
    .Commit '<~ make sure to only commit AT MOST ONCE per transaction.
    
End With 'transaction is rolled back if not committed, connection is closed.
```

When working with an `IUnitOfWork`, best practices would be:

 - **DO** hold the object reference in a `With` block. (e.g. `With New UnitOfWork.FromConnectionString(...)`)
 - **DO** have an active `On Error` statement to graciously handle any errors.
 - **DO** commit or rollback the transaction explicitly in the scope that owns the `IUnitOfWork` object.
 - **AVOID** passing `IUnitOfWork` as a parameter to another object or procedure.
 - **AVOID** accidentally re-entering a `With` block from an error-handling subroutine (i.e. avoid `Resume` or `Resume Next`). If there was an error, execution jumped out of the `With` block that held the references, and the transaction is rolled back and the connection is closed. There's nothing left to clean up.

The scope that creates the `IUnitOfWork` is responsible for knowing what to do with the transaction it encapsulates: we absolutely will be calling other code, but that other code very likely only needs the `IDbCommand` interface, not the whole transaction.

A unit of work can only be committed or rolled back (one of) *once*: keep that responsibility in the scope that owns the `IUnitOfWork`.

### DbConnection

If a transaction isn't needed (recommendation: use one anyway), the API provides the flexibility to use an `IDbConnection` directly. The objects involved are exactly the same as with an `IUnitOfWork`, except now we're going to be creating and injecting them manually (the unit tests demonstrate this).

```vb
With DbConnection.Create(connString)
    'connection is open (no transaction is initiated, but invoking .BeginTransaction would do that)
    
    Dim conn As ADODB.Connection
    Set conn = .AdoConnection '<~ if we want we can access the wrapped connection directly from here, but we don't need to.
    
    
    '...
    
End With 'connection is closed.
```

When working with an `IDbConnection`, best practices would be:
 - **DO** hold the object reference in a `With` block. (e.g. `With DbConnection.Create(...)`)
 - **DO** have an active `On Error` statement to graciously handle any errors.
 - **CONSIDER** passing the `IDbConnection` object to `UnitOfWork.Create`.
 - **CONSIDER** passing the `IDbConnection` object to `DefaultDbCommand.Create`.
 - **AVOID** passing `IDbConnection` as a parameter to another object or procedure.

If an error occurs and execution jumps out of the `With` block (and the `With` block is holding the `IDbConnection` reference), then the connection is already closed when the error handler gets to run.

### IDbCommand

Using an `IUnitOfWork` automatically wires up a factory (`DefaultDbCommandFactory`) that automatically creates these objects for us, but if we want complete flexibility while still leveraging the API, we can wire up the pieces ourselves with classes implementing the `IDbCommand` interface.

There are two implementations:

 - `AutoDbCommand` takes an `IDbConnectionFactory`, owns its database connection, and outputs disconnected recordsets. It also requires an `IDbCommandBase` implementation to be supplied to its `Create` factory method. This class is intended to be used without a `UnitOfWork`.
 - `DefaultDbCommand` takes an `IDbConnection` and `IDbCommandBase` dependencies, and does not own the connection it works with: that's useful when different instances/commands need to run in the same transaction; `UnitOfWork` uses this implementation.
 
#### AutoDbCommand

To use this implementation of `IDbCommand`, you need to inject an abstract factory that's responsible for creating database connections, as well as an `IDbCommandBase` object, which is ultimately responsible for wiring up all the components needed to issue a parameterized ADODB command.

The abstract factory is implemented by the `DbConnectionFactory` class, but the `IDbCommandBase` object needs its own dependencies injected - that's the `IParameterProvider`, an object that's responsible for creating `ADODB.Parameter` objects out of user-provided values:

```vb
Dim mappings As ITypeMap
Set mappings = AdoTypeMappings.Default
'if we want to tweak the ADODB parameter types associated with an intrinsic VBA data type, we can:
''mappings.Mapping("Date") = adVarChar

Dim provider As IParameterProvider
Set provider = AdoParameterProvider.Create(mappings)

Dim baseCommand As IDbCommandBase
Set baseCommand = DbCommandBase.Create(provider)

Dim factory As IDbConnectionFactory
Set factory = New DbConnectionFactory 'the only other implementation is StubDbConnectionFactory, for unit tests.

Dim cmd As IDbCommand
Set cmd = AutoDbCommand.Create(connString, factory, baseCommand)

Dim results As ADODB.Recordset
'simply use '?' ordinal parameters in the command string, and then provide a value for each '?' in the SQL:
Set results = cmd.Execute("SELECT * FROM Table1 WHERE Field1 = ? AND Field2 = ?", 42, "Test")
```

When using an `AutoDbCommand`, we don't need to worry about the database connection: the object uses the supplied factory to create it with the connection string we provide - the command object can be reused for multiple successive calls to various methods, but keep in mind that a new database connection will be created every time.

#### DefaultDbCommand

When more than a single command needs to be sent to a database, it's usually best to run them all using the same connection: you would use a `DefaultDbCommand` for that. The setup is similar to that of `AutoDbCommand`, except no connection factory is involved now - instead, we are working with an `IDbConnection`:

```vb
Dim mappings As ITypeMap
Set mappings = AdoTypeMappings.Default
'if we want to tweak the ADODB parameter types associated with an intrinsic VBA data type, we can:
''mappings.Mapping("Date") = adVarChar

Dim provider As IParameterProvider
Set provider = AdoParameterProvider.Create(mappings)

Dim baseCommand As IDbCommandBase
Set baseCommand = DbCommandBase.Create(provider)

With DbConnection.Create(connString)
    
    Dim cmd As IDbCommand
    Set cmd = DefaultDbCommand.Create(.Self, baseCommand)

    Dim results As ADODB.Recordset
    'simply use '?' ordinal parameters in the command string, and then provide a value for each '?' in the SQL:
    Set results = cmd.Execute("SELECT * FROM Table1 WHERE Field1 = ? AND Field2 = ?", 42, "Test")
    
    '...
    
End With
```

This already covers most scenarios, but there's more!

---

### Advanced & Hybrid Scenarios

#### NULL Parameters

Mapping VBA intrinsic data types to `ADODB.DataTypeEnum` values works fine, ...until you need to supply a `null` value, or supply an *output* parameter to a stored procedure. In such scenarios, you will need to crack the abstraction open and use the underlying ADODB API.

To create an `ADODB.Parameter` with a `NULL` value, you can use the `dataTypeName` optional parameter of the `IParameterProvider.FromValue` method to supply the name of the VBA intrinsic data type. The type of the parameter will be as per the mappings for that intrinsic data type:

```vb
Dim provider As IParameterProvider
Set provider = AdoParameterProvider.Create(mappings)

Dim nullStringValue As ADODB.Parameter
Set nullStringValue = provider.FromValue(Null, dataTypeName:="String") '<~ with default mappings, that's a NULL adVarWChar parameter.
```

The ADODB parameters can then be supplied to the `IDbCommand.ExecuteWithParameters` method, which will append them to the internal ADODB command:

```vb
    Dim cmd As IDbCommand
    Set cmd = DefaultDbCommand.Create(db, baseCommand)
    
    Dim results As ADODB.Recordset
    Set results = cmd.ExecuteWithParameters("SELECT * FROM Table1 WHERE Field1 = ISNULL(?, Field1)", nullStringValue)
```

Note that the SQL command string must be written in such a way that a `NULL` parameter value is still valid a SQL statement: keep in mind that in SQL a `WHERE` clause would use the `IS` operator to determine whether a value is `NULL`.

#### Output Parameters

The ADODB parameters created by the `AdoParameterProvider` can be modified after they are created and before they are used in a command:

```vb
Dim provider As IParameterProvider
Set provider = AdoParameterProvider.Create(mappings)

Dim outParam As ADODB.Parameter
Set outParam = provider.FromValue(42)
outParam.Direction = adParamOutput
```

The modified parameter(s) can then be supplied to `IDbCommand.ExecuteWithParameters` to execute a stored procedure taking output parameters.

#### Stored Procedures

There is no direct support for stored procedures for now, because `Execute` and `ExecuteNonQuery` cover these grounds already, at least as far as SQL Server is concerned - given an `exec` statement:

```vb
Const sql As String = "exec SomeStoredProcedure ?, ?"
AutoDbCommand.Create(connString, New DbConnectionFactory, baseCommand).ExecuteNonQuery sql, 42, "test"
```

That said the `IDbCommand` interface could very well be modified to expose an `ExecuteStoredProcedure` method that takes the name of a stored procedure and its arguments, and wires up an `adCmdStoredProc` ADODB command.

---

