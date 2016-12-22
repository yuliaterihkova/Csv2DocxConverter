package csv2docxconverter;

import com.opencsv.CSVReader;
import java.awt.FileDialog;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.List;
import javax.swing.JOptionPane;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Class encapsulating Main form generation and event handling logic
 * @author Yulia Terikhova
 */
public class MainForm extends javax.swing.JFrame {

    /**
     * Creates new form MainForm
     */
    public MainForm() {
        initComponents();
    }
    
    /**
    *  Columns represented in csv file   
    */ 
    private final String[] columns = new String[]{ "First Name", "Last Name", "Email Address", "Password", "Secondary Email", "Mobile Phone", "Department"};
   
    private String inputFile;
    private String outputFile;
    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        FileTypeButtonGroup = new javax.swing.ButtonGroup();
        InputTextField = new javax.swing.JTextField();
        ChooseInputButton = new javax.swing.JButton();
        OutputTextField = new javax.swing.JTextField();
        ChooseOuputButton = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        GenerateButton = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("CSV2DOCX");
        setBackground(new java.awt.Color(255, 255, 255));
        setResizable(false);

        InputTextField.setEditable(false);
        InputTextField.setText("Choose file location...");
        InputTextField.setToolTipText("");

        ChooseInputButton.setText("Choose");
        ChooseInputButton.setToolTipText("");
        ChooseInputButton.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        ChooseInputButton.setMargin(new java.awt.Insets(2, 10, 2, 10));
        ChooseInputButton.setMaximumSize(new java.awt.Dimension(75, 23));
        ChooseInputButton.setMinimumSize(new java.awt.Dimension(75, 23));
        ChooseInputButton.setPreferredSize(new java.awt.Dimension(80, 23));
        ChooseInputButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ChooseInputButtonActionPerformed(evt);
            }
        });

        OutputTextField.setEditable(false);
        OutputTextField.setText("Choose file location...");
        OutputTextField.setToolTipText("");

        ChooseOuputButton.setText("Choose");
        ChooseOuputButton.setToolTipText("");
        ChooseOuputButton.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        ChooseOuputButton.setMargin(new java.awt.Insets(2, 10, 2, 10));
        ChooseOuputButton.setMaximumSize(new java.awt.Dimension(75, 23));
        ChooseOuputButton.setMinimumSize(new java.awt.Dimension(75, 23));
        ChooseOuputButton.setPreferredSize(new java.awt.Dimension(80, 23));
        ChooseOuputButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ChooseOuputButtonActionPerformed(evt);
            }
        });

        jLabel1.setText("Input CSV file");

        jLabel2.setText("Output DOC\\DOCX file");

        GenerateButton.setText("Generate");
        GenerateButton.setToolTipText("");
        GenerateButton.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        GenerateButton.setMargin(new java.awt.Insets(2, 10, 2, 10));
        GenerateButton.setMaximumSize(new java.awt.Dimension(75, 23));
        GenerateButton.setMinimumSize(new java.awt.Dimension(75, 23));
        GenerateButton.setPreferredSize(new java.awt.Dimension(80, 23));
        GenerateButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                GenerateButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel1)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(InputTextField, javax.swing.GroupLayout.PREFERRED_SIZE, 211, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(6, 6, 6)
                        .addComponent(ChooseInputButton, javax.swing.GroupLayout.PREFERRED_SIZE, 77, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(jLabel2)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(OutputTextField, javax.swing.GroupLayout.PREFERRED_SIZE, 211, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(6, 6, 6)
                        .addComponent(ChooseOuputButton, javax.swing.GroupLayout.PREFERRED_SIZE, 77, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(GenerateButton, javax.swing.GroupLayout.PREFERRED_SIZE, 294, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1)
                .addGap(6, 6, 6)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(1, 1, 1)
                        .addComponent(InputTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(ChooseInputButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(11, 11, 11)
                .addComponent(jLabel2)
                .addGap(9, 9, 9)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(OutputTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(ChooseOuputButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(GenerateButton, javax.swing.GroupLayout.PREFERRED_SIZE, 38, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents
   
    /**
     * Choose input file event handler
     */
    private void ChooseInputButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ChooseInputButtonActionPerformed
        FileDialog dialog = new FileDialog(this, "Choose a file", FileDialog.LOAD);
        dialog.setDirectory("C:\\");
        dialog.setFile("*.csv");
        dialog.setVisible(true);
        
        String filename = dialog.getDirectory() + "\\" + dialog.getFile();
        if (filename == null){
            InputTextField.setText("Choose file location...");
        }
        else{
            inputFile = filename;
            InputTextField.setText(inputFile);
        }
    }//GEN-LAST:event_ChooseInputButtonActionPerformed
    /**
     * Choose output file event handler
     */
    private void ChooseOuputButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ChooseOuputButtonActionPerformed
        FileDialog dialog = new FileDialog(this, "Choose a file", FileDialog.SAVE);
        dialog.setDirectory("C:\\");
        dialog.setFile("*.docx");
        dialog.setVisible(true);
        
        String filename = dialog.getDirectory() + "\\" + dialog.getFile();
        if (filename == null){
            OutputTextField.setText("Choose file location...");
        }
        else{
            outputFile = filename;
            OutputTextField.setText(outputFile);
        }
    }//GEN-LAST:event_ChooseOuputButtonActionPerformed
    /**
     * Convert file event handler
     */
    private void GenerateButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_GenerateButtonActionPerformed
        List content = null;    
    
        try {
            // read data from cvs file
            CSVReader reader = new CSVReader(new FileReader(inputFile));
            content = reader.readAll();
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
        
        try { 
            DocumentGenerator generator = new DocumentGenerator();
            // generate DocX document
            XWPFDocument document = generator.generateDocx(columns, content);
            // save document to file
            document.write(new FileOutputStream(new File(outputFile)));
            document.close();
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }//GEN-LAST:event_GenerateButtonActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MainForm().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton ChooseInputButton;
    private javax.swing.JButton ChooseOuputButton;
    private javax.swing.ButtonGroup FileTypeButtonGroup;
    private javax.swing.JButton GenerateButton;
    private javax.swing.JTextField InputTextField;
    private javax.swing.JTextField OutputTextField;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    // End of variables declaration//GEN-END:variables
}
