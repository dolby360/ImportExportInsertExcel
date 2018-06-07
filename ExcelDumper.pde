class ExcelDumper{ 

  SXSSFWorkbook swb;
  Sheet sh;
  InputStream inp;
  Workbook wb;
  String path;
  Boolean LogsOn;
  
  ExcelDumper(String _path){
    this.path =    _path;
    this.swb       = null;
    this.sh        = null;
    this.inp       = null;
    this.wb        = null;
    this.LogsOn    = false;
  }
  void setLogsOn(boolean value){LogsOn = value;}
  
  Boolean addNewColumn(String columnHeaderName){
    String[][] oldTable = importExcel();
    int newColumnSize = getNumberOfColumn(oldTable[0]) + 1;
    String[][] newTable = new String[oldTable.length][newColumnSize];
    
    //Copy all table
    for(int i = 0; i < oldTable.length; i++){
      for(int j = 0; j < oldTable[i].length; j++){
        newTable[i][j] = oldTable[i][j];
      }
    }
    
    newTable[0][newColumnSize - 1] = columnHeaderName;
    for(int i = 1; i < newTable.length; i++){
      newTable[i][newColumnSize - 1] = null;
    }
    if(exportExcel(newTable)){
      return true;
    }
    return false;     //<>//
  }
  
  String[][] importExcel(){
    String[][] temp;
    String[][] TableWithoutNulls;
    try {
      inp = new FileInputStream(this.path);
      wb = WorkbookFactory.create(inp);
    }catch(Exception e) {
      return null;
    }
    Sheet sheet = wb.getSheetAt(0);
    int sizeX = sheet.getLastRowNum();
    //int sizeX = rwoNumbers;
    int sizeY = 100;
    for (int i=0;i<sizeX;++i) {
      Row row = sheet.getRow(i);
      for (int j=0;j<sizeY;++j) {
        try {
          Cell cell = row.getCell(j);
        }
        catch(Exception e) {
          if (j>sizeY) {
            sizeY = j;
          }
        }
      }
    }
    temp = new String[sizeX][sizeY];
    for (int i=0;i<sizeX;++i) {
      for (int j=0;j<sizeY;++j) {
        Row row = sheet.getRow(i);
        try {
          Cell cell = row.getCell(j);
          if (cell.getCellType()==0 || cell.getCellType()==2 || cell.getCellType()==3)cell.setCellType(1);
          temp[i][j] = cell.getStringCellValue();
        }
        catch(Exception e){
        }
      }
    }
    if(this.LogsOn){ println("Excel file imported: " + this.path + " successfully!"); }
    
    TableWithoutNulls = new String[temp.length][getNumberOfColumn(temp[0])];
    for(int i = 0; i < temp.length; i++){
      for(int j = 0; j < temp[i].length; j++){
        if(temp[i][j] != null){
          TableWithoutNulls[i][j] = temp[i][j];
        }
      }
    }
    return TableWithoutNulls;
  }
  
  boolean insertToExcel(String[] data){
    String temp[][] = importExcel();
    if(temp == null){
      return false;
    }
    String[][] newData = new String[temp.length + 2][getNumberOfColumn(data)];
  
    for(int i = 0; i < temp.length; i++){
      for(int j = 0; j < newData[i].length; j++){
        if(temp[i][j] == null){
          break;
        }
        //println(temp[i][j]);
        newData[i][j] = temp[i][j];
      }
    }
    newData[temp.length] = data;
  
    exportExcel(newData);
    return true;
  }
  
  int getNumberOfColumn(String[] line){
    int counter = 0;
    for(counter = 0; counter < line.length; counter++){
      if(line[counter] == null){
        return counter;
      }
    }
    return counter;
  }
  
  boolean exportExcel(String[] Header,String[][] data){
    for(int i = 0; i < data.length ; i++){
      if(Header.length != getNumberOfColumn(data[i])){
        if(this.LogsOn){ println("Data and header length have to be the same"); } 
        return false;
      }
    }

    
    String newData[][] = new String[data.length + 2][Header.length];
    int i = 0;
    int j = 0;
    int k = 1;
    
    for(; i < Header.length; i++){
      newData[0][i] = Header[i];
      //println(newData[0][i]);
    }
    
    for(i = 0; i < data.length;i++){
      newData[k++] = data[i];
    }
    
    return exportExcel(newData);
  }
  
  boolean exportExcel(String[][] data) {
    SXSSFWorkbook wwb = new SXSSFWorkbook(100);
    Sheet sh = wwb.createSheet();
    int sizeX = data.length;
    int sizeY = data[0].length;
  
    for (int i = 0 ;i < sizeX; ++i) {
      Row row = sh.createRow(i);
      for (int j = 0; j < sizeY; ++j) {
        Cell cell = row.createCell(j);
        if (cell.getCellType()==0 || cell.getCellType()==2 || cell.getCellType()==3)
        cell.setCellType(1);
        cell.setCellValue(data[i][j]);
      }
    }
    try {
      FileOutputStream out = new FileOutputStream(this.path);
      wwb.write(out);
      if(this.LogsOn){ println("Excel file exported: " + this.path + " sucessfully!"); }
      return true;
    }
    catch (Exception e) {
      if(this.LogsOn){ println("Error in saving the file...sorry!"); }
      return false;
    }
  }
  
}
