import javax.swing.*;
import javax.swing.event.HyperlinkEvent;
import javax.swing.event.HyperlinkListener;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;

public class AboutDialog extends JDialog {
    private JPanel contentPane;
    private JButton buttonOK;
    private JEditorPane editorPane1;

    public AboutDialog() {

        initAbout();
        setContentPane(contentPane);
        setModal(true);
        this.setTitle("About");
        this.pack();

        getRootPane().setDefaultButton(buttonOK);

        buttonOK.addActionListener(e -> onOK());

    }

    private void initAbout(){
        editorPane1.setContentType("text/html");//set content as html
        editorPane1.setText("Icons made by <a href=\"http://www.freepik.com\" title=\"Freepik\">Freepik</a>"
                +" from <a href=\"http://www.flaticon.com\" title=\"Flaticon\">www.flaticon.com</a>"
                +" is licensed by <a href=\"http://creativecommons.org/licenses/by/3.0/\" title=\"Creative Commons BY 3.0\" target=\"_blank\">CC 3.0 BY</a>");

        editorPane1.setEditable(false);
        editorPane1.setOpaque(false);

        editorPane1.addHyperlinkListener(hle -> {
            if (HyperlinkEvent.EventType.ACTIVATED.equals(hle.getEventType())) {
                Desktop desktop = Desktop.getDesktop();
                try {
                    desktop.browse(hle.getURL().toURI());
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        });
    }

    private void onOK() {
        // add your code here
        dispose();
    }
}
