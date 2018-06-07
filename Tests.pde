String[] header = {"header","header","header","header","header","header","header","header","header","header" };
String[][] pass = {{"1","1","1","1","1","1","1","1","1","1"},{"2","2","2","2","2","2","2","2","2","2"},
                   {"3","3","3","3","3","3","3","3","3","3"},{"4","4","4","4","4","4","4","4","4","4"},
                   {"5","5","5","5","5","5","5","5","5","5"},{"6","6","6","6","6","6","6","6","6","6"},
                   {"7","7","7","7","7","7","7","7","7","7"}};
String[][] faile = { {"1"} };

void executeTests(){
  
  if(!test_exportExcel()){
    println("========================");
    println("Export excel test failed");
    println("========================");
    return;
  }
  println("**********************");
  println("Export excel test pass");
  println("**********************");
  
  
  
  if(!test_importExcel()){
    println("========================");
    println("Import excel test failed");
    println("========================");
    return;
  }
  println("**********************");
  println("Import excel test pass");
  println("**********************");
  
  
  
  
  if(!test_insertToExcel()){
    println("===========================");
    println("insert to excel test failed");
    println("===========================");
    return;
  }
  println("*************************");
  println("Insert to excel test pass");
  println("*************************");
  
  
  
  
  
  if(!test_insertNewColumnToExcel()){
    println("===================================");
    println("Insert new column to excel test failed");
    println("===================================");
    return;
  }
  println("*********************************");
  println("Insert new row to excel test pass");
  println("*********************************");
}


boolean test_insertNewColumnToExcel(){
  String[][] myTable;

  ExcelDumper ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
  ed.setLogsOn(true);
  
  ed.addNewColumn("New!!!!!");
  myTable = ed.importExcel();
  if(myTable[0][myTable[0].length - 1].equals("New!!!!!")){ //<>//
    return true;
  }
  return false;
}

      
boolean test_importExcel(){
  String[][] compare = {
                   {"header","header","header","header","header","header","header","header","header","header"},
                   {"1","1","1","1","1","1","1","1","1","1"},{"2","2","2","2","2","2","2","2","2","2"},
                   {"3","3","3","3","3","3","3","3","3","3"},{"4","4","4","4","4","4","4","4","4","4"},
                   {"5","5","5","5","5","5","5","5","5","5"},{"6","6","6","6","6","6","6","6","6","6"},
                   {"7","7","7","7","7","7","7","7","7","7"}};
  String[][] myTable;
  
  
  ExcelDumper ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
  ed.setLogsOn(true);
  myTable = ed.importExcel();
  
  
  for(int i = 0; i < myTable.length; i++){
    for(int j = 0; j < myTable[i].length; j++){
      if(myTable[i][j].equals(compare[i][j]) == false){
        return false;
      }
    }
  }
  
  return true;
}

boolean test_exportExcel(){

  ExcelDumper ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
  ed.setLogsOn(true);
  if(ed.exportExcel(header,faile)){
    //Shuold not return true 
    //The length of the array of data has to be the same as the header.
    return false;
  }
  
  ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
  ed.setLogsOn(true);
  if(ed.exportExcel(header,pass)){
    return true;
  }else{
    return false;
  }

}

boolean test_insertToExcel(){
  String[][] compare = {
                   {"header","header","header","header","header","header","header","header","header","header"},
                   {"1","1","1","1","1","1","1","1","1","1"},{"2","2","2","2","2","2","2","2","2","2"},
                   {"3","3","3","3","3","3","3","3","3","3"},{"4","4","4","4","4","4","4","4","4","4"},
                   {"5","5","5","5","5","5","5","5","5","5"},{"6","6","6","6","6","6","6","6","6","6"},
                   {"7","7","7","7","7","7","7","7","7","7"},{"8","8","8","8","8","8","8","8","8","8"}};
  String[][] myTable;
  String[] toInsert = {"8","8","8","8","8","8","8","8","8","8"};
  
  ExcelDumper ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
  ed.setLogsOn(true);
  ed.insertToExcel(toInsert);

  //Checking if the column inserted OK
  myTable = ed.importExcel();
  
  for(int i = 0; i < myTable.length; i++){
    for(int j = 0; j < myTable[i].length; j++){
      if(myTable[i][j].equals(compare[i][j]) == false){
        return false;
      }
    }
  }
  return true;
}
