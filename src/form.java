import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;


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


    public form(){
        setContentPane(contentPanel);

        buttonOK.addActionListener(actionEvent -> onAdd());
        ButtonEDIT.addActionListener(actionEvent -> onEdit());
        buttonExit.addActionListener(actionEvent -> onExit());
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


