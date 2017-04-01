import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;


/**
 * Created by terpal on 4/1/17.
 */
public class form extends JFrame {
    private JPanel contentPanel;
    private JTable table1;
    private JTextField SEARCHTextField;
    private JButton buttonOK;
    private JButton ButtonEDIT;
    private JButton buttonExit;
    private JButton buttonImport;
    private JFileChooser jChooser;
    private String path;

    public form(){
        setContentPane(contentPanel);

        buttonOK.addActionListener(actionEvent -> onAdd());
        ButtonEDIT.addActionListener(actionEvent -> onEdit());
        buttonExit.addActionListener(actionEvent -> onExit());



        readExcel();
        buttonImport.addActionListener(actionEvent -> {
            jChooser = new JFileChooser();
            jChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
            jChooser.addChoosableFileFilter(new FileNameExtensionFilter("MS Office Documents", "docx", "xlsx", "pptx"));
            jChooser.showOpenDialog(null);

            File file = jChooser.getSelectedFile();
            if (file == null){

            }else {
                if(!file.getName().endsWith("xls")){
                    JOptionPane.showMessageDialog(null,
                            "Please select only Excel file.",
                            "Error",JOptionPane.ERROR_MESSAGE);
                }else {
                    path = file.getAbsolutePath();
                }

            }
        });
    }

    private void readExcel(){

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


    public static void main(String[] args) {
        EventQueue.invokeLater(() -> {
            form ex = new form();

            ex.pack();
            ex.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            ex.setVisible(true);
        });

    }

}


