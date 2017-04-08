/**
 * Created by terpal on 4/1/17.
 */

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.awt.*;
import java.io.*;
import java.util.Vector;


public class form extends JFrame {
    private JPanel contentPanel;
    private JTable table;
    //private JTextField SEARCHTextField;
    private JButton buttonSave;
    private JButton ButtonEDIT;
    private JButton buttonExit;
    private JButton buttonImport;
    private JButton buttonDelete;
    private JTextField SEARCHONTITLETextField;
    private JButton searchButton;
    private JLabel RowCount;
    private JFileChooser jChooser;
    private String path;

    // contains table Column
    static Vector headers = new Vector();
    // contains Data from Excel File
    static Vector data = new Vector();
    // Model is used to construct
    DefaultTableModel model = null;
    static int tableWidth = 0;
    static int tableHeight = 0;


    public form() throws IOException {
        setContentPane(contentPanel);

        buttonSave.addActionListener(actionEvent -> onSave());
        ButtonEDIT.addActionListener(actionEvent -> onEdit());
        buttonExit.addActionListener(actionEvent -> onExit());
        buttonImport.addActionListener(actionEvent -> onImport());
        buttonDelete.addActionListener(actionEvent -> onDelete());

        fillDefaultList();
    }

    /**
     * Import Button
     * User choose a excel file to import.
     */
    private void onImport(){
        jChooser = new JFileChooser();
        jChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        jChooser.addChoosableFileFilter(new FileNameExtensionFilter("MS Office Documents", "xls", "xlsx"));
        jChooser.showOpenDialog(null);

        File file = jChooser.getSelectedFile();
        if (file == null){

        }else {
            path = file.getAbsolutePath();
            fillData(file);
            model = new DefaultTableModel(data, headers);
            table.setModel(model);
            updateCountText();
        }
    }

    /**
     * Save Button
     */
    private void onSave(){
        new WorkbookFactory();
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        Row headerRow = sheet.createRow(0);
        Row row = sheet.createRow(2);
        TableModel model = table.getModel();

        for(int headings = 0; headings < model.getColumnCount(); headings++){
            headerRow.createCell(headings).setCellValue(model.getColumnName(headings));
        }

        for(int rows = 0; rows < model.getRowCount() ; rows++){
            for(int cols = 0; cols < table.getColumnCount(); cols++){
                try{
                    row.createCell(cols).setCellValue(model.getValueAt(rows, cols).toString());
                }catch (NullPointerException e){
                    row.createCell(cols).setCellValue("");
                }
            }

            row = sheet.createRow((rows + 3));
        }
        try {
            wb.write(new FileOutputStream("watch.xls"));//Save the file
            JOptionPane.showMessageDialog(null,
                    "The List was saved",
                    "SAVE",JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Exit Button
     */
    private void onExit(){
        dispose();
    }

    /**
     * Edit Button
     */
    private void onEdit(){
        DefaultTableModel model = (DefaultTableModel) table.getModel();
        model.addRow(new Object[]{"Column 1", "Column 2", "Column 3", "Column 4"});
        updateCountText();
    }

    /**
     * Delete Button
     */
    private void onDelete(){
        DefaultTableModel model = (DefaultTableModel) table.getModel();
        int deleteRow = table.getSelectedRow();
        model.removeRow(deleteRow);
        updateCountText();
    }

    /**
     * Opens the Default excel file
     */
    private void fillDefaultList(){
        File defFile = null;
        defFile = new File("watch.xls");
        if(defFile.exists() && !defFile.isDirectory()) {
            fillData(defFile);
        }else{
            createNewWorkbook();
        }

        model = new DefaultTableModel(data, headers);
        //tableWidth = model.getColumnCount() * 150;
        //tableHeight = model.getRowCount() * 25;
        //table.setPreferredSize(new Dimension( tableWidth, tableHeight));
        table.setModel(model);
        updateCountText();
    }

    /**
     * Imports the data from the excel file to the JTable
     *
     * @param file
     */
    private void fillData(File file) {

        HSSFWorkbook workbook = null;
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            //createNewWorkbook();
            e.printStackTrace();
        }
        if (file.getAbsolutePath().endsWith("xls")) {
            // HSSFWorkbook, InputStream
            try {
                NPOIFSFileSystem fs = new NPOIFSFileSystem(inputStream);
                workbook = new HSSFWorkbook(fs.getRoot(), true);
                fs.close();
            } catch (IOException e) {
                e.printStackTrace();
            }

        } else {
            //throw new IllegalArgumentException("The specified file is not Excel file");
            JOptionPane.showMessageDialog(null,
                    "Please select only Excel file. (.xls)",
                    "Error",JOptionPane.ERROR_MESSAGE);
        }

        HSSFSheet sheet = workbook.getSheetAt(0);
        HSSFRow row = sheet.getRow(0);

        headers.clear();
        for (int i = 0; i < 4; i++)
        {
            HSSFCell cell1 = row.getCell(i);
            headers.add(cell1.toString());
        }

        data.clear();
        for (int j = 1; j < sheet.getLastRowNum() + 1; j++)
        {
            Vector d = new Vector();
            row=sheet.getRow(j);
            if (isRowEmpty(row) == false){
                int noofrows=row.getLastCellNum();
                for (int i = 0; i < noofrows; i++)
                {    //To handle empty excel cells
                    HSSFCell cell=row.getCell(i,
                            org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK );
                    d.add(cell.toString());
                }
                d.add("\n");
                data.add(d);

            }
        }


    }

    private void createNewWorkbook(){
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow((short)0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Anime");
        Cell cell1 = row.createCell(1);
        cell1.setCellValue("Status");
        Cell cell2 = row.createCell(2);
        cell2.setCellValue("Comments");
        Cell cell3 = row.createCell(3);
        cell3.setCellValue("Rate");

        // Write the output to a file
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("watch.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        File startFile = null;
        startFile = new File("watch.xls");
        fillData(startFile);
    }

    private void updateCountText(){
        DefaultTableModel model = (DefaultTableModel) table.getModel();
        RowCount.setText("Num: " + String.valueOf(model.getRowCount()));
    }

    /**
     * Checks if a row is Empty
     *
     * @param row
     * @return
     */
    public static boolean isRowEmpty(Row row) {
        if (row == null){
            return true;
        }
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
                return false;
        }
        return true;
    }

    /**
     * Get the workbook of a file
     *
     * @param excelFilePath
     * @return
     * @throws IOException
     */
    private Workbook getWorkbook(String excelFilePath)
            throws IOException {

        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = null;

        if (excelFilePath.endsWith("xlsx")) {

            OPCPackage pkg = null;
            try {
                pkg = OPCPackage.open(inputStream);
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
            workbook = new XSSFWorkbook(pkg);
            pkg.close();

        } else if (excelFilePath.endsWith("xls")) {
            // HSSFWorkbook, InputStream
            NPOIFSFileSystem fs = new NPOIFSFileSystem(inputStream);
            workbook = new HSSFWorkbook(fs.getRoot(), true);
            fs.close();

        } else {
            //throw new IllegalArgumentException("The specified file is not Excel file");
            JOptionPane.showMessageDialog(null,
                    "Please select only Excel file.",
                    "Error",JOptionPane.ERROR_MESSAGE);
        }

        return workbook;
    }

    public static void main(String[] args) {
        EventQueue.invokeLater(() -> {
            form ex = null;
            try {
                ex = new form();
            } catch (IOException e) {
                e.printStackTrace();
            }

            ex.pack();
            ex.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            ex.setVisible(true);
        });

    }

}


