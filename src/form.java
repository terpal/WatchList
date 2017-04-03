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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Vector;



public class form extends JFrame {
    private JPanel contentPanel;
    private JTable table;
    private JTextField SEARCHTextField;
    private JButton buttonOK;
    private JButton ButtonEDIT;
    private JButton buttonExit;
    private JButton buttonImport;
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

        buttonOK.addActionListener(actionEvent -> onAdd());
        ButtonEDIT.addActionListener(actionEvent -> onEdit());
        buttonExit.addActionListener(actionEvent -> onExit());
        buttonImport.addActionListener(actionEvent -> onImport());
    }


    public void onImport(){
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
            tableWidth = model.getColumnCount() * 150;
            tableHeight = model.getRowCount() * 25;
            table.setPreferredSize(new Dimension( tableWidth, tableHeight));
            table.setModel(model);
        }
    }

    public void onAdd(){
        dispose();
    }

    public void onExit(){
        dispose();
    }

    public void onEdit(){
        dispose();
    }

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

    void fillData(File file) {

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
        HSSFRow row=sheet.getRow(0);

        headers.clear();
        for (int i = 0; i < row.getLastCellNum(); i++)
        {
            HSSFCell cell1 = row.getCell(i);
            headers.add(cell1.toString());
        }

        data.clear();
        for (int j = 1; j < sheet.getLastRowNum() + 1; j++)
        {
            Vector d = new Vector();
            row=sheet.getRow(j);
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


