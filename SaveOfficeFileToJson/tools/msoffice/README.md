# SmallTools
This are some classes containing some static methods targeting at saving MicroSoft Office file to json file.



*How To Use It:* 

1. It has two dependencies, one of which is Apache POI, another is Gson. You can add their Maven dependencies into pom.xml as follows:

```
		<dependency>
			<groupId>com.google.code.gson</groupId>
			<artifactId>gson</artifactId>
			<version>2.6.2</version>
   	</dependency>
   	<dependency>
   		<groupId>org.apache.poi</groupId>
   		<artifactId>poi-ooxml</artifactId>
   		<version>3.14</version>
   	</dependency>
```

2. After Importing it, you can call *XlsxTransfer.readFile* to get data from xlsx file. There are 3 arguments for this static function: 

   1. *String filePath* which is the xlsx file path you want to read. 

   2. *HashMap\<String, Integer> columnMap* which denotes what columns you want to get and their matched column number, respectively, if you want to read all columns, just make it null.

   3. *int sheetNumber* denotes which sheet you want to read. 

   4. Note: this function returns a *ArrayList\<HashMap\<String, String>>*.

3. After reading it, you can call *XlsxTransfer.saveAsJsonFile* to save it as a json file. There are 2 arguments for this static function:

   1. *ArrayList\<HashMap\<String, String>> array* is what you got from *readFile* function.

   2. *String fileName* is the path where you want to save the generated json file.

   3. Note: If the file path you denoted exists a file with same name already, it will throw an IOException with console message:"File Exists!".

4. Both of these 2 functions throws IOExceptions.

5. Will update more such functions based on these 2 libraries, probably.