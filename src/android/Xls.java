package net.claudiomedeiros.xls;


import android.util.Log;
 
import org.apache.cordova.CallbackContext;
import org.apache.cordova.CordovaPlugin;
import org.json.JSONObject;
import org.json.JSONArray;
import org.json.JSONException;

import android.widget.Toast;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;

import android.os.Environment;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import jxl.*;
import jxl.format.Alignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class Xls extends CordovaPlugin {
    public static final String ACTION_SAVE_XLS = "saveXLS";
     
    @Override
    public boolean execute(String action, JSONArray args, CallbackContext callbackContext) throws JSONException {
 
        try {
            if(ACTION_SAVE_XLS.equals(action)){
                /*  Modelo dos parâmetros
                    var data = [
                                    {id:"1", name:"claudio"} ,
                                    {id:"2", name:"marta"} ,
                                    {id:"3", name:"isabela"} 
                               ];
                    dirname = "ExcelAPI";
                    filename = "file-example.xls";
                    sheetname = "Plan1";
                */
                
                JSONObject params = args.getJSONObject(0);

                try{
		   //Define o nome do diretório
                    this.dirname = params.getString("dirname");
                    
                    //Define o nome do arquivo
                    WritableWorkbook wb = this.createWorkbook(params.getString("filename"));

                    //Define o nome da planilha
                    WritableSheet sheetObject = this.createSheet(wb, params.getString("sheetname"), 1);
                    
                    
                    JSONArray lineItems = params.getJSONArray("data");
                    
                    //Loop pelas linhas de dados
                    for (int i = 0, size = lineItems.length(); i < size; i++){
                        JSONObject objectInArray = lineItems.getJSONObject(i);
                        jsonObjectToCell(sheetObject, objectInArray);
                    }
                    
                    //Escrevendo os dados no arquivo
                    wb.write();
                    
                    //Fechando o arquivo
                    wb.close();
                    
                    //Precisa adicionar esse callback para informar ao phonegap que
                    //a execução ocorreu com sucesso
                    callbackContext.success(params);
                    return true;
                }
                catch(IOException ex){
                    Log.e(TAG, ex.getStackTrace().toString());
                    Log.e(TAG, ex.getMessage());
                }
            }
           
            callbackContext.error("Invalid action");
            return false;
        }
        catch(Exception e) {
            System.err.println("Exception: " + e.getMessage());
            callbackContext.error(e.getMessage());
            Log.v(TAG, e.getStackTrace().toString());
            return false;
        } 
    }     
    
    public String dirname;
    
   /**
     * @property int rowPosition
     * Inicia a contagem das linhas em 1, para que a linha 0 seja 
     */
    public int rowPosition = 1;
    
   /**
     * @property Boolean hasTitles
     * Armazena a informação se a planilha já tem títulos definidos
     */
    public Boolean hasTitles = false;
    
   /**
     * @param  sheetObj - Sheet object
     * @para   obj      - JSON object
     * @return void
     */
    public void jsonObjectToCell(WritableSheet sheetObj, JSONObject obj){
    	try{
    		//Inicia a contagem das colunas em 0
    		int columnPosition = 0;
            
            Iterator<?> keys = obj.keys();
	        while( keys.hasNext() ){
	        	//Pega o nome da chave
	        	String key = (String)keys.next();
	        	
                //Pega o valor da chave
	            String value = obj.getString(key);
	            
	            try{
	            	//Se não tem título definido
		        	if(!hasTitles){
		        		//Escreve os títulos na primeira linha
		        		this.writeCell(columnPosition, 0, key, true, sheetObj);
		        	}
		        	
                    //Escreve os valores dos objetos na célula a partir da segunda linha (rowPosition=1)
	        		this.writeCell(columnPosition, rowPosition, value, false, sheetObj);
		            
	        		Log.i(TAG, key +":"+ value);
		            
                    //Adiciona 1 ao contador de colunas
		            columnPosition++;
	            }
	            catch(WriteException we){
	            	Log.i(TAG, we.toString());
	            }
	        }
            this.rowPosition++;
            this.hasTitles = true;

    	}
    	catch(JSONException e){
    		Log.i(TAG, e.toString());
    	}
    }
    
  /**
    * @param fileName - the name to give the new workbook file
    * @return - a new WritableWorkbook with the given fileName 
    */
    public WritableWorkbook createWorkbook(String fileName){
        //exports must use a temp file while writing to avoid memory hogging
        WorkbookSettings wbSettings = new WorkbookSettings(); 				
        wbSettings.setUseTemporaryFileDuringWrite(true);   
     
        //get the sdcard's directory
        File sdCard = Environment.getExternalStorageDirectory();
        //add on the your app's path
        File dir = new File(sdCard.getAbsolutePath() + "/" + this.dirname);
        //make them in case they're not there
        dir.mkdirs();
        //create a standard java.io.File object for the Workbook to use
        File wbfile = new File(dir, fileName);
     
        WritableWorkbook wb = null;
     
        try{
        //create a new WritableWorkbook using the java.io.File and
        //WorkbookSettings from above
        wb = Workbook.createWorkbook(wbfile,wbSettings); 
        }catch(IOException ex){
            Log.e(TAG,ex.getStackTrace().toString());
            Log.e(TAG, ex.getMessage());
        }
     
        return wb;	                
    }
 
    String TAG = "xls";
    
  /**
    * @param wb - WritableWorkbook to create new sheet in
    * @param sheetName - name to be given to new sheet
    * @param sheetIndex - position in sheet tabs at bottom of workbook
    * @return - a new WritableSheet in given WritableWorkbook
    */
    public WritableSheet createSheet(WritableWorkbook wb, String sheetName, int sheetIndex){
       //create a new WritableSheet and return it
       return wb.createSheet(sheetName, sheetIndex);
    }

   /**
     * @param columnPosition - column to place new cell in
     * @param rowPosition - row to place new cell in
     * @param contents - string value to place in cell
     * @param headerCell - whether to give this cell special formatting
     * @param sheet - WritableSheet to place cell in
     * @throws RowsExceededException - thrown if adding cell exceeds .xls row limit
     * @throws WriteException - Idunno, might be thrown
     */
    public void writeCell(int columnPosition, int rowPosition, String contents, boolean headerCell,
        WritableSheet sheet) throws RowsExceededException, WriteException{
        //create a new cell with contents at position
        Label newCell = new Label(columnPosition,rowPosition,contents);
     
        if (headerCell){
            //give header cells size 10 Arial bolded 	
            WritableFont headerFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            WritableCellFormat headerFormat = new WritableCellFormat(headerFont);
            //center align the cells' contents
            headerFormat.setAlignment(Alignment.CENTRE);
            newCell.setCellFormat(headerFormat);
        }
     
        sheet.addCell(newCell);
    }
}
