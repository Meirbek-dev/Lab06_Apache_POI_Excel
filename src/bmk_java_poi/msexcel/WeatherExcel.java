package bmk_java_poi.msexcel;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class WeatherExcel extends javax.swing.JFrame {

    private static final long serialVersionUID = 1L;

    class Launch_Thread extends Thread {

        @Override
        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");
            try {
                modifData(dir + "weather_template.xls", dir + "weather.xls",
                        jTextField_Date.getText(),
                        jTextField_Day.getText(),
                        jTextField_Time.getText());
                Desktop.getDesktop().open(new File(dir + "weather.xls"));
            } catch (IOException ex) {
                System.err.println("Error modifData!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    // Метод создания отчета
    private void modifData(String inputFileName, String outputFileName, String Date, String Day, String Time) throws IOException {

        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFileName)));
        HSSFSheet sheet = wb.getSheetAt(0);
        sheet.getRow(2).getCell(2).setCellValue(Date);
        sheet.getRow(3).getCell(2).setCellValue(Day);
        sheet.getRow(4).getCell(2).setCellValue(Time);

        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }

    public WeatherExcel() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTextField_Date = new javax.swing.JTextField();
        jTextField_Day = new javax.swing.JTextField();
        jTextField_Time = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Прогноз погоды в MS Excel");
        setLocation(new java.awt.Point(1000, 1000));
        setResizable(false);
        getContentPane().setLayout(null);

        jTextField_Date.setCaretColor(new java.awt.Color(255, 0, 255));
        getContentPane().add(jTextField_Date);
        jTextField_Date.setBounds(140, 34, 210, 26);

        jTextField_Day.setCaretColor(new java.awt.Color(255, 0, 255));
        jTextField_Day.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_DayActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Day);
        jTextField_Day.setBounds(140, 60, 210, 35);

        jTextField_Time.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_TimeActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Time);
        jTextField_Time.setBounds(140, 95, 210, 25);

        jButton1.setFont(new java.awt.Font("Times New Roman", 0, 14)); // NOI18N
        jButton1.setText("Экспортировать в Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(70, 370, 190, 30);

        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/bmk_java_poi/msexcel/forecast.PNG"))); // NOI18N
        getContentPane().add(jLabel1);
        jLabel1.setBounds(0, 0, 350, 411);

        setBounds(0, 0, 365, 447);
    }// </editor-fold>//GEN-END:initComponents

    private void jTextField_TimeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_TimeActionPerformed
    }//GEN-LAST:event_jTextField_TimeActionPerformed

    private void jTextField_DayActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_DayActionPerformed
    }//GEN-LAST:event_jTextField_DayActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new Launch_Thread().start();
    }//GEN-LAST:event_jButton1ActionPerformed

    public static void main(String args[]) {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(WeatherExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        java.awt.EventQueue.invokeLater(() -> {
            new WeatherExcel().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JTextField jTextField_Date;
    private javax.swing.JTextField jTextField_Day;
    private javax.swing.JTextField jTextField_Time;
    // End of variables declaration//GEN-END:variables
}
