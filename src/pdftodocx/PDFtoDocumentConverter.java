
package pdftodocx;

import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class PDFtoDocumentConverter extends javax.swing.JFrame {


    public PDFtoDocumentConverter() {
        initComponents();

    }


public static void pdfToDocxConverter(String pdfFilePath, String docxFilePath) throws IOException {
        try (InputStream inputStream = new FileInputStream(new File(pdfFilePath));
            BufferedInputStream bufferedInputStream =new BufferedInputStream(inputStream);
             PDDocument document = PDDocument.load(bufferedInputStream);
             XWPFDocument doc = new XWPFDocument()) {

            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);
            
            XWPFParagraph paragraph = doc.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText(text);
        

            try (FileOutputStream out = new FileOutputStream(new File(docxFilePath+".docx"));
            BufferedOutputStream bufferedOutputStream=new BufferedOutputStream(out)){
                doc.write(bufferedOutputStream);
            
            }
        }
    }

    public static void convertPDFToPPTX(String pdfFilePath, String pptxFilePath) throws IOException {

        try (InputStream inputStream = new FileInputStream(new File(pdfFilePath));
             PDDocument document = PDDocument.load(inputStream);
             XMLSlideShow ppt = new XMLSlideShow()) {

            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);

                // Create a new slide
                XSLFSlide slide = ppt.createSlide();

                // Create a text box and add text to the slide
                XSLFTextShape textBox = slide.createTextBox();
                textBox.setText(text);

                // Adjust text box position and size as needed
                textBox.setAnchor(new java.awt.Rectangle(50, 50, 500, 300));
            
            try (FileOutputStream out = new FileOutputStream(new File(pptxFilePath+".pptx"))) {
                ppt.write(out);
            }
        }
    }

    // @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        slctPDfLbl = new javax.swing.JLabel();
        darkCheckBox = new javax.swing.JCheckBox();
        slctPdfTextField = new javax.swing.JTextField();
        fileBtn = new javax.swing.JButton();
        slctOutputLbl = new javax.swing.JLabel();
        slctOutputTextField = new javax.swing.JTextField();
        formatComboBox = new javax.swing.JComboBox<>();
        covertBtn = new javax.swing.JButton();
        aboutBtn = new javax.swing.JButton();
        supportBtn = new javax.swing.JButton();
        closeBtn = new javax.swing.JButton();
        folderBtn = new javax.swing.JButton();
        formatLbl = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("PDF-to-Docment Convertor");
        setLocation(new java.awt.Point(400, 100));
        setResizable(false);
        ImageIcon image =new ImageIcon("C:\\Users\\bahir\\OneDrive\\Documents\\NetBeansProjects\\PDFtoDOCX\\src\\pdftodocx\\photo_2024-02-21_13-16-11.jpg");
        setIconImage(image.getImage());

        jPanel1.setBackground(new java.awt.Color(255, 255, 255));

        slctPDfLbl.setText("Select PDF File");

        darkCheckBox.setText("Night Mode");
        darkCheckBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                darkCheckBoxActionPerformed(evt);
            }
        });

        slctPdfTextField.setText("C:\\Users\\ ");

        fileBtn.setText("Files");
        fileBtn.setFocusable(false);
        fileBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                fileBtnActionPerformed(evt);
            }
        });

        slctOutputLbl.setText("Select Output Folder ");

        slctOutputTextField.setText("C:\\Users\\");
        slctOutputTextField.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                slctOutputTextFieldActionPerformed(evt);
            }
        });

        formatComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { ".docx", ".pptx", " " }));

        covertBtn.setText("Convert");
        covertBtn.setFocusable(false);
        covertBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                covertBtnActionPerformed(evt);
            }
        });

        aboutBtn.setText("About");
        aboutBtn.setFocusable(false);
        aboutBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                aboutBtnActionPerformed(evt);
            }
        });

        supportBtn.setText("Support");
        supportBtn.setFocusable(false);
        supportBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                supportBtnActionPerformed(evt);
            }
        });

        closeBtn.setText("Close");
        closeBtn.setFocusPainted(false);
        closeBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                closeBtnActionPerformed(evt);
            }
        });

        folderBtn.setText("Folder");
        folderBtn.setFocusable(false);
        folderBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                folderBtnActionPerformed(evt);
            }
        });

        formatLbl.setText("Format ");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(slctPdfTextField, javax.swing.GroupLayout.DEFAULT_SIZE, 305, Short.MAX_VALUE)
                            .addComponent(slctOutputTextField))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 7, Short.MAX_VALUE)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(fileBtn, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(folderBtn, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addContainerGap())
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(slctPDfLbl, javax.swing.GroupLayout.PREFERRED_SIZE, 129, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(darkCheckBox, javax.swing.GroupLayout.PREFERRED_SIZE, 98, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(slctOutputLbl, javax.swing.GroupLayout.PREFERRED_SIZE, 130, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(formatLbl, javax.swing.GroupLayout.PREFERRED_SIZE, 47, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(formatComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 260, Short.MAX_VALUE))))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(aboutBtn, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(supportBtn, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(closeBtn))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(covertBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(slctPDfLbl)
                    .addComponent(darkCheckBox, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(slctPdfTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(fileBtn))
                .addGap(18, 18, 18)
                .addComponent(slctOutputLbl)
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(slctOutputTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(folderBtn))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(formatComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(formatLbl))
                .addGap(28, 28, 28)
                .addComponent(covertBtn)
                .addGap(304, 304, 304)
                .addComponent(aboutBtn)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(supportBtn)
                    .addComponent(closeBtn))
                .addGap(13, 13, 13))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void covertBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed

        String pdfPath=slctPdfTextField.getText().toString();
        String docxPath=slctOutputTextField.getText().toString();
        String docxcombox=".docx",pptxcombox=".pptx";
        // int progres=0;
        // JProgressBar progress=new JProgressBar(0,100);
        //     progress.setStringPainted(true);
        //     progress.setString("Converting "+progres +"%");
        //     progress.setBounds(120,220,250,25);
        //     progress.setValue(progres);
        //     add(progress);

        if(formatComboBox.getSelectedItem().toString()== docxcombox){

            try{
                     pdfToDocxConverter(pdfPath,docxPath); 
                    JOptionPane.showMessageDialog(this, "The pdf file is successfully converted to docement file", "Sucessful", JOptionPane.INFORMATION_MESSAGE);
            }catch(IOException e){
                JOptionPane.showMessageDialog(this,"The  File is not found\nPlease set pdf file path \nand Set Output path and output file name", "Error", JOptionPane.INFORMATION_MESSAGE);

            }
        }
        if(formatComboBox.getSelectedItem().toString()==pptxcombox){
        
            try{
                convertPDFToPPTX(pdfPath,docxPath); 
                JOptionPane.showMessageDialog(this, "The pdf file is successfully converted to PowerPoint(pptx) file", "Successful", JOptionPane.INFORMATION_MESSAGE);
            }catch(IOException e){
                JOptionPane.showMessageDialog(this,"The  File is not found\nPlease set pdf file path \nand Set Output path and output file name \n", "Error", JOptionPane.INFORMATION_MESSAGE);

            }
        }


    }//GEN-LAST:event_jButton2ActionPerformed

    /**
     * @param evt
     */
    private void slctOutputTextFieldActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_slctOutputTextFieldActionPerformed
        
    }//GEN-LAST:event_slctOutputTextFieldActionPerformed

    private void darkCheckBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_darkCheckBoxActionPerformed

        if(darkCheckBox.isSelected()){
            this.setBackground(new java.awt.Color(51, 51, 51));
            jPanel1.setBackground(new java.awt.Color(51, 51, 51)); 
            slctPDfLbl.setForeground(new java.awt.Color(255, 255, 255));
            slctOutputLbl.setForeground(new java.awt.Color(255, 255, 255));
            formatLbl.setForeground(new java.awt.Color(255, 255, 255));

        }
        else if(!(darkCheckBox.isSelected())){
            jPanel1.setBackground(new java.awt.Color(255, 255, 255));
            slctPDfLbl.setForeground(new java.awt.Color(51, 51, 51));
            slctOutputLbl.setForeground(new java.awt.Color(51, 51, 51));
            formatLbl.setForeground(new java.awt.Color(51, 51, 51));

        }


    }//GEN-LAST:event_darkCheckBoxActionPerformed

    private void fileBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        JFileChooser fileChooser =new JFileChooser("c:");
        fileChooser.showOpenDialog(null);
       String filepath = fileChooser.getSelectedFile().toString();
       slctPdfTextField.setText(filepath);


    }//GEN-LAST:event_jButton1ActionPerformed

    private void folderBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_folderBtnActionPerformed
        JFileChooser fileChooser =new JFileChooser("c:");
        fileChooser.showSaveDialog(null);
        String folderpath=fileChooser.getSelectedFile().toString();
        slctOutputTextField.setText(folderpath);


        // TODO add your handling code here:
    }//GEN-LAST:event_folderBtnActionPerformed

    private void aboutBtnActionPerformed(java.awt.event.ActionEvent evt) {                                         
        JTextArea textArea=new JTextArea();
        textArea.setEditable(false);
        textArea.setText("This tool is used to convert pdf document to word document(docx)"+
                         "\nand used to convert pdf to PowerPoint(pptx) \nthis tool is specially useful for students"+
                        "\nThis tool is prepared by Pogrammer Bahiru Walelign"+
                         "\nI am computer Science student and also \nI have experience some programming languages like c++,java...."+
                        "\n\nAddress\nContact :- +251914370474"+
                         "\nTelegeram :- t.me/Bahir_21"+
                         "\nEmail :- bahiruwalelign@gmail.com\n"
                         );
    
        JOptionPane.showMessageDialog(this, textArea, "About", JOptionPane.INFORMATION_MESSAGE);


    }                                          

    private void supportBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_supportBtnActionPerformed
        JTextArea textArea = new JTextArea();
        textArea.setEditable(false);
        textArea.setText("If you want any support about application \nand other any question you can ask below Telegram address\n\nTelegram :- t.me/Bahir_21\n\n");
        JOptionPane.showMessageDialog(this , textArea,"Help",JOptionPane.INFORMATION_MESSAGE);

        // TODO add your handling code here:
    }//GEN-LAST:event_supportBtnActionPerformed

    private void closeBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_closeBtnActionPerformed
        System.exit(ABORT);

        // TODO add your handling code here:
    }//GEN-LAST:event_closeBtnActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
      
        
//         new PDFtoDocumentConverter().setVisible(true);
        
        java.awt.EventQueue.invokeLater(new Runnable() {
            @Override
            public void run() {
                new PDFtoDocumentConverter().setVisible(true);
            }
        });
        
        
        
        
        
        
        
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton aboutBtn;
    private javax.swing.JButton closeBtn;
    private javax.swing.JButton covertBtn;
    private javax.swing.JCheckBox darkCheckBox;
    private javax.swing.JButton fileBtn;
    private javax.swing.JButton folderBtn;
    private javax.swing.JComboBox<String> formatComboBox;
    private javax.swing.JLabel formatLbl;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JLabel slctOutputLbl;
    private javax.swing.JTextField slctOutputTextField;
    private javax.swing.JLabel slctPDfLbl;
    private javax.swing.JTextField slctPdfTextField;
    private javax.swing.JButton supportBtn;
    // End of variables declaration//GEN-END:variables
}
