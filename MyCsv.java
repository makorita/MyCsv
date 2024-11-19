import java.io.*;
import java.util.*;
import java.util.regex.*;

import java.nio.file.*;
import java.nio.charset.*;

import org.apache.poi.*;
import org.apache.poi.ss.formula.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.xssf.usermodel.*;

public class MyCsv{
	public static final String OUTPUT_COLNAME="出力名";
	public static final String COL_DEFNAME="col#";
	
	private ArrayList<String> header;
	private ArrayList<HashMap<String,String>> content;
	
	public MyCsv(){
		header=new ArrayList<String>();
		content=new ArrayList<HashMap<String,String>>();
	}
	
	//↓↓↓getter,setter
	public ArrayList<String> getHeader(){
		return header;
	}
	
	public void setHeader(ArrayList<String> header){
		this.header=header;
	}
	
	public void addHeader(String colName){
		header.add(colName);
	}
	
	public void removeHeader(String colName){
		header.remove(header.indexOf(colName));
	}
	
	public ArrayList<HashMap<String,String>> getContent(){
		return content;
	}
	
	public int getContentSize(){
		return content.size();
	}
	
	public void setContent(ArrayList<HashMap<String,String>> content){
		this.content=content;
	}
	
	public void addContent(ArrayList<HashMap<String,String>> srcContent){
		for(HashMap<String,String> curMap:srcContent){
			content.add(curMap);
		}
	}
	
	public HashMap<String,String> getRow(int index){
		if(content.size()<=index)return null;
		
		return content.get(index);
	}
	
	public void addRow(HashMap<String,String> curRow){
		content.add(curRow);
	}
	
	public String getValue(int rowIndex,String colName){
		if(rowIndex>content.size()-1)return null;
		if(!content.get(rowIndex).containsKey(colName))return null;
		
		return content.get(rowIndex).get(colName);
	}
	
	public void setValue(int rowIndex,String colName,String value){
		if(rowIndex>content.size()-1)return;
		
		content.get(rowIndex).put(colName,value);
	}
	
	public void replaceValue(String colName,String befStr,String aftStr){
		for(HashMap<String,String> curMap:content){
			if(curMap.get(colName)==null)continue;
			curMap.put(colName,curMap.get(colName).replaceAll(befStr,aftStr));
		}
	}
	
	public void renameColName(String befCol,String aftCol){
		int index=header.indexOf(befCol);
		header.set(index,aftCol);
		
		for(HashMap<String,String> curMap:content){
			if(!curMap.containsKey(befCol))continue;
			String tmpValue=curMap.get(befCol);
			curMap.remove(befCol);
			curMap.put(aftCol,tmpValue);
		}
	}
	
	public void setDefaultValue(String colName,String defaultValue){
		for(HashMap<String,String> curMap:content){
			if(curMap.containsKey(defaultValue))continue;
			
			curMap.put(colName,defaultValue);
		}
	}
	
	//↓↓↓loader,saver,outputer
	public void loadDataCsv(ArrayList<String> rowStrList){
		for(String curStr:rowStrList){
			String[] word=curStr.split(",");
			HashMap<String,String> curMap=new HashMap<String,String>();
			content.add(curMap);
			for(int i=0;i<word.length;i++){
				if(i>header.size()-1)header.add(i,COL_DEFNAME+i);
				if(word[i].length()>0)curMap.put(header.get(i),word[i]);
			}
		}
	}
	
	public void loadCsv(File srcFile) throws Exception{
		BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(srcFile), "SHIFT-JIS"));
		String tmpStr=br.readLine();
		String[] word=tmpStr.split(",");
		setHeader(new ArrayList<String>(Arrays.asList(word)));
		
		String line;
		ArrayList<String> rowStrList=new ArrayList<String>();
		while ((line = br.readLine()) != null) {
			rowStrList.add(line);
		}
		br.close();
		
		loadDataCsv(rowStrList);
	}
	
	public void loadCsv(String srcPath) throws Exception{
		File tmpFile=new File(srcPath);
		loadCsv(tmpFile);
	}
	
	public void loadTabFile(File srcFile,boolean headerFlag) throws Exception{
		BufferedReader br = new BufferedReader(new FileReader(srcFile));
		if(headerFlag){
			String tmpStr=br.readLine();
			String[] word=tmpStr.split("\t");
			setHeader(new ArrayList<String>(Arrays.asList(word)));
		}
		
		String line;
		ArrayList<String> rowStrList=new ArrayList<String>();
		while ((line = br.readLine()) != null) {
			rowStrList.add(line.replaceAll("\t",","));
		}
		
		br.close();
		
		loadDataCsv(rowStrList);
	}
	
	public void loadTabFile(String srcPath,boolean headerFlag) throws Exception{
		File tmpFile=new File(srcPath);
		loadTabFile(tmpFile,headerFlag);
	}
	
	public void loadExcel(File srcFile,String sheetName) throws Exception{
		Workbook wb=null;
		Sheet sheet=null;
		wb=WorkbookFactory.create(new FileInputStream(srcFile));
		sheet=wb.getSheet(sheetName);
		for(int rowIndex=0;rowIndex<=sheet.getLastRowNum();rowIndex++){
			Row row=sheet.getRow(rowIndex);
			if(rowIndex==0 && row==null)break;
			if(row==null)continue;
			
			if(rowIndex==0 && row.getZeroHeight())break;
			if(row.getZeroHeight())continue;
			
			if(rowIndex==0){
				ArrayList<String> tmpList=new ArrayList<String>();
				for(int cellIndex=0;cellIndex<row.getLastCellNum();cellIndex++){
					Cell cell=row.getCell(cellIndex);
					String tmpStr=cell.getStringCellValue();
					tmpList.add(tmpStr);
				}
				setHeader(tmpList);
				continue;
			}
			
			HashMap<String,String> tmpMap=new HashMap<String,String>();
			addRow(tmpMap);
			for(int cellIndex=0;cellIndex<row.getLastCellNum();cellIndex++){
				Cell cell=row.getCell(cellIndex);
				if(cell==null)continue;
				String curCellStr=null;
				try{
					curCellStr=cell.getStringCellValue();
				}catch(IllegalStateException e){
					curCellStr=String.format("%.0f",cell.getNumericCellValue());
				}
				tmpMap.put(getHeader().get(cellIndex),curCellStr);
			}
		}
		
		wb.close();
	}
	
	public void loadExcel(String srcPath,String sheetName) throws Exception{
		File tmpFile=new File(srcPath);
		loadExcel(tmpFile,sheetName);
	}
	
	public void saveCsv(String dstPath) throws Exception{
		PrintWriter wr = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(dstPath),"Shift-JIS")));
		
		for(int i=0;i<header.size();i++){
			if(i==0)wr.print(header.get(i));
			else wr.print(","+header.get(i));
		}
		wr.println();
		
		for(HashMap<String,String> curRow:content){
			for(int i=0;i<header.size();i++){
				if(i==0){
					if(curRow.containsKey(header.get(i)))wr.print(curRow.get(header.get(i)));
				}else{
					if(curRow.containsKey(header.get(i)))wr.print(","+curRow.get(header.get(i)));
					else wr.print(",");
				}
			}
			wr.println();
		}
		
		wr.close();
	}
	
	public void outputGroupreplace(String templatePath,String dstDir) throws Exception{
		List<String> lines=null;
		Path path=Paths.get(templatePath);
		lines=Files.readAllLines(path,StandardCharsets.UTF_8);
		
		for(HashMap<String,String> curMap:content){
			if(!curMap.containsKey(OUTPUT_COLNAME))continue;
			File curFile=new File(dstDir+"/"+curMap.get(OUTPUT_COLNAME));
			PrintWriter wr=null;
			if(curFile.exists())wr=new PrintWriter(new FileWriter(curFile,true));
			else wr=new PrintWriter(new FileWriter(curFile));
			
			for(String curStr:lines){
				for(String headerStr:header){
					if(curMap.containsKey(headerStr))curStr=curStr.replaceAll("<"+headerStr+">",curMap.get(headerStr));
					else curStr=curStr.replaceAll("<"+headerStr+">","");
				}
				wr.println(curStr);
			}
			
			wr.close();
		}
	}
	
	//↓↓↓transformer
	public ArrayList<String> getCol(String colName){
		ArrayList<String> returnList=new ArrayList<String>();
		for(HashMap<String,String> curMap:content){
			if(curMap.containsKey(colName))returnList.add(curMap.get(colName));
			else returnList.add(null);
		}
		
		return returnList;
	}
	
	public HashMap<String,String> getMapping(String keyCol,String valueCol){
		HashMap<String,String> returnMap=new HashMap<String,String>();
		
		for(HashMap<String,String> curMap:content){
			if(!curMap.containsKey(keyCol))continue;
			if(!curMap.containsKey(valueCol))continue;
			if(returnMap.containsKey(curMap.get(keyCol)))continue;
			
			returnMap.put(curMap.get(keyCol),curMap.get(valueCol));
		}
		
		return returnMap;
	}
	
	public MyCsv mapMerge(String keyCol,String valueCol,HashMap<String,String> srcMap){
		MyCsv returnCsv=new MyCsv();
		
		ArrayList<String> tmpHeader=new ArrayList<String>(getHeader());
		tmpHeader.add(valueCol);
		returnCsv.setHeader(tmpHeader);
		
		for(HashMap<String,String> curMap:content){
			HashMap<String,String> tmpMap=new HashMap<String,String>(curMap);
			returnCsv.addRow(tmpMap);
			
			if(!tmpMap.containsKey(keyCol))continue;
			if(!srcMap.containsKey(tmpMap.get(keyCol)))continue;
			
			tmpMap.put(valueCol,srcMap.get(tmpMap.get(keyCol)));
		}
		
		return returnCsv;
	}
	
	public static MyCsv csvMerge(File rootDir) throws Exception{
		MyCsv returnCsv=null;
		
		File[] fileList=rootDir.listFiles();
		for(File curFile:fileList){
			MyCsv curCsv=new MyCsv();
			curCsv.loadCsv(curFile);
			
			if(returnCsv==null)returnCsv=curCsv;
			else{
				returnCsv.addContent(curCsv.getContent());
			}
		}
		
		return returnCsv;
	}
	
	public MyCsv listReplace(String keyCol,String value,ArrayList<String> replaceList){
		MyCsv returnCsv=new MyCsv();
		returnCsv.setHeader(new ArrayList<String>(getHeader()));
		
		for(HashMap<String,String> curMap:content){
			if(curMap.containsKey(keyCol) && curMap.get(keyCol).equals(value)){
				for(String curStr:replaceList){
					HashMap<String,String> tmpMap=new HashMap<String,String>(curMap);
					tmpMap.put(keyCol,curStr);
					returnCsv.addRow(tmpMap);
				}
			}else returnCsv.addRow(curMap);
		}
		
		return returnCsv;
	}
	
	public MyCsv sort(String keyCol,boolean ascendFlag){
		MyCsv returnCsv=new MyCsv();
		returnCsv.setHeader(new ArrayList<String>(getHeader()));
		
		for(HashMap<String,String> curMap:content){
			returnCsv.addRow(curMap);
		}
		
		Collections.sort(returnCsv.getContent(), new Comparator<HashMap<String, String>>() {
			@Override
			public int compare(HashMap<String, String> map1, HashMap<String, String> map2) {
				String value1 = map1.get(keyCol);
				String value2 = map2.get(keyCol);
				if(ascendFlag)return value1.compareTo(value2);
				else return value2.compareTo(value1);
			}
		});
		
		return returnCsv;
	}
	
	//↓↓↓filter
	public MyCsv clone(){
		MyCsv returnCsv=new MyCsv();
		returnCsv.setHeader(new ArrayList<String>(getHeader()));
		
		for(HashMap<String,String> curMap:content){
			HashMap<String,String> tmpMap=new HashMap<String,String>(curMap);
			returnCsv.addRow(tmpMap);
		}
		
		return returnCsv;
	}
	public MyCsv reFilter(String keyCol,String reStr,boolean matchFlag){
		MyCsv returnCsv=new MyCsv();
		returnCsv.setHeader(new ArrayList<String>(getHeader()));
		
		for(HashMap<String,String> curMap:content){
			if(matchFlag){
				if(!curMap.containsKey(keyCol))continue;
				
				Pattern p=Pattern.compile(reStr);
				Matcher m=p.matcher(curMap.get(keyCol));
				if(m.find())returnCsv.addRow(curMap);
			}else{
				if(!curMap.containsKey(keyCol)){
					returnCsv.addRow(curMap);
					continue;
				}
				
				Pattern p=Pattern.compile(reStr);
				Matcher m=p.matcher(curMap.get(keyCol));
				if(!m.find())returnCsv.addRow(curMap);
			}
		}
		
		return returnCsv;
	}
	
	public MyCsv containsIpFilter(String keyCol,String checkedStr,boolean matchFlag) throws Exception{
		MyCsv returnCsv=new MyCsv();
		returnCsv.setHeader(new ArrayList<String>(getHeader()));
		
		for(HashMap<String,String> curMap:content){
			if(matchFlag){
				if(!curMap.containsKey(keyCol))continue;
				if(!Address.isLegalIP(curMap.get(keyCol)))continue;
				
				Address checkedAddr=new Address(checkedStr);
				Address checkAddr=new Address(curMap.get(keyCol));
				if(checkAddr.containsAddress(checkedAddr))returnCsv.addRow(curMap);
				
			}else{
				if(!curMap.containsKey(keyCol)){
					returnCsv.addRow(curMap);
					continue;
				}
				if(!Address.isLegalIP(curMap.get(keyCol))){
					returnCsv.addRow(curMap);
					continue;
				}
				
				Address checkedAddr=new Address(checkedStr);
				Address checkAddr=new Address(curMap.get(keyCol));
				if(!checkAddr.containsAddress(checkedAddr))returnCsv.addRow(curMap);
				
			}
		}
		
		return returnCsv;
	}
	
	public MyCsv containedIpFilter(String keyCol,String checkStr,boolean matchFlag) throws Exception{
		MyCsv returnCsv=new MyCsv();
		returnCsv.setHeader(new ArrayList<String>(getHeader()));
		
		for(HashMap<String,String> curMap:content){
			if(matchFlag){
				if(!curMap.containsKey(keyCol))continue;
				if(!Address.isLegalIP(curMap.get(keyCol)))continue;
				
				Address checkedAddr=new Address(curMap.get(keyCol));
				Address checkAddr=new Address(checkStr);
				if(checkAddr.containsAddress(checkedAddr))returnCsv.addRow(curMap);
				
			}else{
				if(!curMap.containsKey(keyCol)){
					returnCsv.addRow(curMap);
					continue;
				}
				if(!Address.isLegalIP(curMap.get(keyCol))){
					returnCsv.addRow(curMap);
					continue;
				}
				
				Address checkedAddr=new Address(curMap.get(keyCol));
				Address checkAddr=new Address(checkStr);
				if(!checkAddr.containsAddress(checkedAddr))returnCsv.addRow(curMap);
				
			}
		}
		
		return returnCsv;
	}
	
	//↓↓↓utility
	public void showAll(){
		for(String curStr:header){
			System.out.print(curStr+",");
		}
		
		System.out.println();
		for(HashMap<String,String> curMap:content){
			for(String curStr:header){
				System.out.print(curMap.get(curStr)+",");
			}
			System.out.println();
		}
	}
	
	public static void deleteFileInFolder(String dirPath) throws Exception{
		File rootDir=new File(dirPath);
		File[] fileList=rootDir.listFiles();
		for(File curFile:fileList)curFile.delete();
	}
}
