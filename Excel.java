package org.firstinspires.ftc.teamcode;

import android.os.Environment;

import java.io.File;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.LinkedList;
import java.util.List;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Excel {
    private static final int DEFAULT_SHEET = 0;

    private static int writeXLS( String path, List<ArrayList<Double>> table, ArrayList<String> title) {
        File file = createXLS(path);

        if ( file==null ) {
            return CREATE_FAIL;
        } else {
            return addDouble( file,table, title);
        }
    }

    private static File createXLS(String path) {
        File file = null;

        try {
            file=new File(path);

            WritableWorkbook book = Workbook.createWorkbook(file);
            WritableSheet sheet = book.createSheet("task", DEFAULT_SHEET);

            book.write();
            book.close();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return file;
        }
    }

    private static int addDouble( File file, List<ArrayList<Double>> table ,ArrayList<String> title){
        try {
            Workbook wb = Workbook.getWorkbook(file);

            WritableWorkbook book= Workbook.createWorkbook(file,wb);
            WritableSheet sheet = book.getSheet(DEFAULT_SHEET);

            if(!title.isEmpty()) {
                for (int attr = 0; attr < title.size(); attr++) {
                    Label label = new Label(attr, 0, title.get(attr));
                    sheet.addCell(label);
                }

                for (int task = 0; task < table.size(); task++) {
                    List<Double> attrs = table.get(task);

                    for (int attr = 0; attr < attrs.size(); attr++) {
                        jxl.write.Number number = new jxl.write.Number(attr, task + 1, attrs.get(attr));
                        sheet.addCell(number);
                    }
                }
            }
            else{
                for (int task = 0; task < table.size(); task++) {
                    List<Double> attrs = table.get(task);

                    for (int attr = 0; attr < attrs.size(); attr++) {
                        jxl.write.Number number = new jxl.write.Number(attr, task, attrs.get(attr));
                        sheet.addCell(number);
                    }
                }
            }
            book.write();
            book.close();

            return 0;
        } catch (Exception e) {
            e.printStackTrace();
            return ADD_DATA_FAIL;
        }
    }

    private static final int CREATE_FAIL = -1;
    private static final int ADD_DATA_FAIL = -2;

    private List<ArrayList<Double>> wholeData = new LinkedList<>();
    private ArrayList<String> title = new ArrayList<>();
    private ArrayList<Double> eachLineData = new ArrayList<>();
    private boolean isSaved = true;
    private String head = null;

    public void addData(double number){
        eachLineData.add(number);
    }

    public void updateNewLine(){
        isSaved = false;
        wholeData.add(eachLineData);
        eachLineData = new ArrayList<>();
    }

    public void saveFile(){
        if(!isSaved){
            Calendar calendar = Calendar.getInstance();
            String fileName =  head +
                    " "+calendar.get(Calendar.MONTH)+
                    "."+calendar.get(Calendar.DAY_OF_MONTH)+
                    " "+calendar.get(Calendar.HOUR_OF_DAY)+
                    ":"+calendar.get(Calendar.MINUTE)+
                    ":"+calendar.get(Calendar.SECOND)+
                    ".xls";
            String sdCardRoot = Environment.getExternalStorageDirectory().getAbsolutePath();
            writeXLS(sdCardRoot + File.separator + "数据" + File.separator + fileName,wholeData,title);
            isSaved = true;
            eachLineData = new ArrayList<>();
            wholeData = new LinkedList<>();
        }
    }
    public Excel(String headOfFileName, ArrayList<String> titleOfExcel){
        head = headOfFileName;
        title = titleOfExcel;
    }
    public Excel(String headOfFileName){
        head = headOfFileName;
    }
}
