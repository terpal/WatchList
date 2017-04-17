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
import javax.swing.event.HyperlinkEvent;
import javax.swing.event.HyperlinkListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Vector;


public class form extends JFrame {
    private JPanel contentPanel;
    private JTable table;
    private JButton buttonSave;
    private JButton ButtonEDIT;
    private JButton buttonExit;
    private JButton buttonDelete;
    private JTextField SEARCHONTITLETextField;
    private JButton searchButton;
    private JLabel RowCount;
    private JFileChooser jChooser;
    private String path;
    private JMenuItem newMenuItem ;
    private JMenuItem importMenuItem ;
    private JMenuItem exportMenuItem ;
    private JMenuItem aboutMenuItem;


    // contains table Column
    static Vector headers = new Vector();
    // contains Data from Excel File
    static Vector data = new Vector();
    // Model is used to construct
    DefaultTableModel model = null;



    public form() throws IOException {
        StoreRestore store = new StoreRestore();
        path = store.Deserialize();
        initUI();
    }

    private void initUI(){
        setContentPane(contentPanel);

        table.setAutoCreateRowSorter(true);

        createMenuBar();
        buttonSave.addActionListener(actionEvent -> onSave());
        ButtonEDIT.addActionListener(actionEvent -> onEdit());
        buttonExit.addActionListener(actionEvent -> onExit());
        //buttonImport.addActionListener(actionEvent -> onImport());
        buttonDelete.addActionListener(actionEvent -> onDelete());
        searchButton.addActionListener(actionEvent -> onSearch());
        importMenuItem.addActionListener(ActionListener -> onImport());
        exportMenuItem.addActionListener(ActionListener -> onExport());
        newMenuItem.addActionListener(ActionListener -> onNew());
        aboutMenuItem.addActionListener(ActionListener -> onAbout());
        SEARCHONTITLETextField.addFocusListener(new FocusAdapter() {
            @Override
            public void focusGained(FocusEvent e) {
                super.focusGained(e);
                SEARCHONTITLETextField.setText("");
            }
        });
        fillDefaultList();
    }

    private void createMenuBar() {

        JMenuBar menuBar = new JMenuBar();

        // Menu File / About
        JMenu fileMenu = new JMenu("File");
        JMenu helpMenu = new JMenu("Help");

        menuBar.add(fileMenu);
        menuBar.add(helpMenu);

        // File->New / Import / Export
        newMenuItem = new JMenuItem("New");
        importMenuItem = new JMenuItem("Import");
        exportMenuItem = new JMenuItem("Export");
        fileMenu.add(newMenuItem);
        fileMenu.add(importMenuItem);
        fileMenu.add(exportMenuItem);

        aboutMenuItem = new JMenuItem("About");
        helpMenu.add(aboutMenuItem);

        //helpMenu.addActionListener(ActionListener -> onAbout());

        setJMenuBar(menuBar);
    }

    private void onAbout(){


        AboutDialog aboutDialog = new AboutDialog();
        aboutDialog.setVisible(true);
    }

    private void onNew(){
        String s = (String)JOptionPane.showInputDialog(
                this,
                "Enter Name of the WatchList",
                "Create New WatchList",
                JOptionPane.INFORMATION_MESSAGE,
                null,
                null,
                null);

        if (s != "" && s!= null){
            s += ".xls";
            createNewWorkbook(s);

            File file = null;
            file = new File(s);
            fillData(file);
        }else {
            System.exit(0);
        }
    }

    private void onSearch(){
        TableModel model = table.getModel();
        final TableRowSorter<TableModel> sorter = new TableRowSorter<TableModel>(model);
        table.setRowSorter(sorter);

        String text = SEARCHONTITLETextField.getText();
        if (text.length() == 0) {
            sorter.setRowFilter(null);
        } else {
            sorter.setRowFilter(RowFilter.regexFilter("(?i)" +  text)); // "(?i)" case insensitive
            sorter.setSortKeys(null);

        }
    }

    /**
     * Import Button
     * User choose a excel file to import.
     */
    private void onImport(){
        jChooser = new JFileChooser();
        jChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        jChooser.setFileFilter(new FileNameExtensionFilter("MS excel file" , "xls"));
        jChooser.showOpenDialog(null);

        File file = jChooser.getSelectedFile();
        if (file == null){

        }else {
            path = file.getAbsolutePath();
            StoreRestore store = new StoreRestore();
            store.Serialize(path);
            fillData(file);
        }
    }

    private void onExport(){
        jChooser = new JFileChooser();
        jChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        //jChooser.getChoosableFileFilters();
        jChooser.setFileFilter(new FileNameExtensionFilter("MS" , "xls"));
        //jChooser.s
        int retval = jChooser.showSaveDialog(exportMenuItem);

        if (retval == JFileChooser.APPROVE_OPTION) {

            File file = jChooser.getSelectedFile();
            path = file.getAbsolutePath();
            if (file != null) {
                if (path.endsWith(".xls")){

                }else{
                    path += ".xls";
                }
                Save(path);
            }
        }

    }

    /**
     * Save Button
     */
    private void onSave(){
        Save(path);
    }

    private void Save(String savePath){
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
            wb.write(new FileOutputStream(savePath));//Save the file
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
        model.addRow(new Object[]{"Title", "Status", "Comments", "Rate"});
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
     * Opens the Default / Last Used excel file
     */
    private void fillDefaultList(){
        String name = path;
        if (name != null){
            File defFile = new File(name);
            if(defFile.exists() && !defFile.isDirectory()) {
                fillData(defFile);
            }else{
                JOptionPane.showMessageDialog(null,
                        "You must create a new WatchList.",
                        "Info",JOptionPane.INFORMATION_MESSAGE);
                onNew();
            }
        }else{
            JOptionPane.showMessageDialog(null,
                    "You must create a new WatchList.",
                    "Info",JOptionPane.INFORMATION_MESSAGE);
            onNew();
        }
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

        model = new DefaultTableModel(data, headers);
        table.setModel(model);
        updateCountText();

        ArrayList dataStatus = new ArrayList();
        int colStatus = 1;     // Status
        dataStatus.add("ongoing");
        dataStatus.add("completed");
        dataStatus.add("to watch");
        dataStatus.add("stopped");
        setComboBox(colStatus,dataStatus);

        ArrayList dataRate = new ArrayList();
        int colRate = 3;     // Rate
        dataRate.addAll(Arrays.asList("1", "2", "3","4","5","6","7","8","9","10"));
        setComboBox(colRate,dataRate);

    }

    private void createNewWorkbook(String name){
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow((short)0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Title");
        Cell cell1 = row.createCell(1);
        cell1.setCellValue("Status");
        Cell cell2 = row.createCell(2);
        cell2.setCellValue("Comments");
        Cell cell3 = row.createCell(3);
        cell3.setCellValue("Rate");

        // Write the output to a file
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(name);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            wb.write(fileOut);
            fileOut.close();
            path = name;
            StoreRestore store = new StoreRestore();
            store.Serialize(path);
        } catch (IOException e) {
            e.printStackTrace();
        }
        File startFile = null;
        startFile = new File(name);
        fillData(startFile);
    }

    private void updateCountText(){
        DefaultTableModel model = (DefaultTableModel) table.getModel();
        RowCount.setText("Count: " + String.valueOf(model.getRowCount()));
    }

    private void setComboBox(int column, ArrayList data){
        TableColumnModel columnModel = table.getColumnModel();
        TableColumn xColumn = columnModel.getColumn(column);
        JComboBox comboBox = new JComboBox();

        for (int i =0 ; i < data.size(); i++){
            comboBox.addItem(data.get(i));
        }

        xColumn.setCellEditor(new DefaultCellEditor(comboBox));
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
     * Saves the Last Used File Path

    private void Serialize(){

        try {
            FileOutputStream fileOut =
                    new FileOutputStream("lastUsedFile.ser");
            ObjectOutputStream out = new ObjectOutputStream(fileOut);
            out.writeObject(path);
            out.close();
            fileOut.close();
        }catch(IOException i) {
            i.printStackTrace();
        }
    }

    /**
     *  Restores the Last Used File Path

    private void Deserialize() {
        try {
            FileInputStream fileIn = null;
            fileIn = new FileInputStream("lastUsedFile.ser");
            ObjectInputStream in = new ObjectInputStream(fileIn);
            path = (String) in.readObject();
            in.close();
            fileIn.close();
        } catch (IOException i) {
            i.printStackTrace();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }
    */

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
            ImageIcon img = new ImageIcon("icons/list32.png");
            ex.setIconImage(img.getImage());
            ex.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            ex.setTitle("WatchList");
            ex.pack();
            ex.setVisible(true);
        });

    }

}


