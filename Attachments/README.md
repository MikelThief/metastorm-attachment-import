The SQL script can be used in conjunction with the .net CLR project. The script may need modifying to work with more / less folders.


The CLR has to be compiled and the dll imported as an assembly to the SQL server. 

In order for the CLR functions to work you need to run the following against the database

    EXEC sp_changedbowner 'sa'
    Create Function WriteToFile(@Input varbinary(max), @path nvarchar(max), @append bit )
       returns nvarchar(max)        
       AS
       External name CLR.[CLR.UserDefinedFunctions].WriteToFile
       
The way the C# works is a bit convoluted so I'll try to explain here. Metastorm apparently does the following when saving attachments:

```
Attachments are sent to the database as MIME-encoded (base64-encoded) strings. 
The component then converts this MIME-encoded string (in unicode) to a byte stream before writing it to the database as a BLOB (Oracle) or Image (SQL Server).
```

This is ambiguous; "unicode" is not actually a text encoding; it is the general system of representing symbols as a number. 

The .Net framework has an encoding called "Unicode", but this is a red herring, this encoding is actually UTF-16.

There are two types of attachment saved. Look at the econtents column and you can see some start like `0x7B00350030003100460032003300350046002D00370`
And some start like `0x7B35303146323335462D373546302D343936342D394`
The latter failed using my original code.

The difference between those two formats is that one of them has 00 bytes in between each of the data bytes. This corresponds to UTF-16, where all 
symbols are 16 bits, aka, 2 bytes. The compact data without those 00 bytes should be plain ASCII.

The 00 bytes actually provide a solution for the problem, because we can check those 00 bytes and use their existence to choose the appropriate 
encoding.

```
if (!reader.Read())
        return new byte[0];
    if (reader[0] == System.DBNull.Value)
        return new byte[0];
    byte[] data = (byte[])reader[0];
    if (data.Length == 0)
        return new byte[0];
    String base64String
    if (data.Length > 1 && data[1] == 00)
        base64String = Encoding.Unicode.GetString(data);
    else
        base64String = Encoding.ASCII.GetString(data);
    // Cuts off the GUID, and takes care of any trailing 00 bytes.
    String truncatedString = base64String.Substring(38).TrimEnd('\0');
    return Convert.FromBase64String(truncatedString);
```