import com.sun.tools.javac.Main;

import javax.swing.*;
import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.event.*;
import java.io.IOException;

public class MainForm {

    public static final int ONE_SECOND = 1000;
    public static final int TWO_SECONDS = 2000;

    public JFrame jFrame;

    private JLabel variationLabel;
    private JButton btnSetVariation;
    private JLabel variationCords;
    private JLabel productCords;
    private JLabel gasketTypeCords;
    private JLabel gasketThicknessCords;
    private JLabel npsCords;
    private JLabel ansiCords;
    private JPanel cutGasketPanel;
    private JLabel gasketIdCords;
    private JLabel gasketOdCords;
    private JLabel numBoltHolesCords;
    private JLabel boltHoleDiaCords;
    private JTextField textFieldAmountOfEntry;
    private JButton btnStart;
    private JButton btnSetProduct;
    private JButton btnSetGasketType;
    private JButton btnSeGasketThickness;
    private JButton btnSetNps;
    private JButton btnSetAnsi;
    private JButton btnSetGasketOd;
    private JButton btnSetGasketId;
    private JButton btnSetBoltHoles;
    private JButton btnSetBoltDiameter;
    private JButton btnSetExcelTitleBar;
    private JButton btnSetWeight;
    private JButton btnSetSku;
    private JButton btnSetPrice;
    private JButton btnSetOk;
    private JLabel weightCords;
    private JLabel skuCords;
    private JLabel priceCords;
    private JLabel okayCords;
    private JLabel excelCords;
    private JLabel excelEditCords;
    private JButton btnSetExcelEditCords;
    private JLabel saveCords;
    private JButton btnSetSave;

    // Cords
    private Point ptVariation;
    private Point ptProduct;
    private Point ptGasketType;
    private Point ptGasketThickness;
    private Point ptNps;
    private Point ptAnsi;
    private Point ptGasketOd;
    private Point ptGasketId;
    private Point ptNumOfBoltHoles;
    private Point ptBoltHoleDia;
    private Point ptExcelTitle;
    private Point ptWeight;
    private Point ptSku;
    private Point ptPrice;
    private Point ptOk;
    private Point ptExcelEdit;
    private Point ptSave;

    // Options
    private String productOptions[] = {
            "Any",
            "KLINGERSIL C-4401",
            "KLINGERSIL C-4430",
            "400MAX",
            "KLINGERtc-1003",
            "KLINGERtc-1005",
            "KLINGERtc-1006"
    };

    private String gasketTypeOptions[] = {
            "Any",
            "Full Face",
            "Ring"
    };

    private String gasketThicknessOptions[] = {
            "Any",
            "1/8\"",
            "1/16\""
    };

    private String npsOptions[] = {
            "Any",
            "1/2\"",
            "3/4\"",
            "1\"",
            "1 1/5\"",
            "1 1/4\"",
            "1 1/2\"",
            "2\"",
            "2 1/2\"",
            "3\"",
            "3 1/2\"",
            "4\"",
            "5\"",
            "6\"",
            "8\"",
            "10\"",
            "12\"",
            "14\"",
            "16\"",
            "18\"",
            "20\"",
            "24\""
    };

    private String AnsiOptions[] = {
            "Any",
            "P1 (150#)",
            "P2 (300#)"
    };

    private String gasketOdOptions[] = {
            "Any",
            "1.88 ",
            "2.12 ",
            "2.25 ",
            "2.62 ",
            "2.88 ",
            "3.00 ",
            "3.25 ",
            "3.38 ",
            "3.50 ",
            "3.75 ",
            "3.88 ",
            "4.12 ",
            "4.25 ",
            "4.38 ",
            "4.62 ",
            "4.63 ",
            "4.88 ",
            "5.00 ",
            "5.12 ",
            "5.25 ",
            "5.38 ",
            "5.88 ",
            "6.00 ",
            "6.12 ",
            "6.38 ",
            "6.50 ",
            "6.88 ",
            "7.00 ",
            "7.12 ",
            "7.50 ",
            "7.75 ",
            "8.25 ",
            "8.50 ",
            "8.75 ",
            "9.00 ",
            "9.88 ",
            "10.00 ",
            "11.00 ",
            "12.12 ",
            "12.50 ",
            "13.38 ",
            "13.50 ",
            "14.25 ",
            "15.00 ",
            "16.00 ",
            "16.13 ",
            "16.62 ",
            "17.50 ",
            "17.75 ",
            "19.00 ",
            "19.12 ",
            "20.25 ",
            "20.50 ",
            "21.00 ",
            "21.25 ",
            "21.62 ",
            "23.00 ",
            "23.50 ",
            "23.88 ",
            "25.00 ",
            "25.50 ",
            "25.75 ",
            "27.50 ",
            "28.00 ",
            "28.25 ",
            "30.50 ",
            "32.00 ",
            "36.00 "
    };

    private String gacketIdOptions[] = {
            "Any",
            "0.84",
            "1.06",
            "1.31",
            "1.66",
            "1.91",
            "2.38",
            "2.88",
            "3.50",
            "4.00",
            "4.50",
            "5.56",
            "6.62",
            "8.62",
            "10.75",
            "12.75",
            "14.00",
            "16.00",
            "18.00",
            "20.00",
            "24.00"
    };

    private String numberOfBoltHolesOptions[] = {
            "Any",
            "4",
            "8",
            "12",
            "16",
            "20",
            "24",
            ""
    };

    private String boltHoleDiameterOptions[] = {
            "Any",
            "0.62",
            "0.75",
            "0.88",
            "1.00",
            "1.12",
            "1.25",
            "1.38",
            "1.62",
            ""
    };

    public static void main(String[] args) {
        new MainForm();
    }

    public MainForm() {
        // Setup Window
        jFrame = new JFrame();
        jFrame.add(cutGasketPanel);
        jFrame.setTitle("Midnet Media - Automatic Data Entry Program");
        jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        jFrame.setFocusable(true);
        jFrame.pack();

        // Setup Buttons

        btnSetVariation.addActionListener(addCordsAfterButtonPress(ptVariation = new Point(), variationCords));
        btnSetProduct.addActionListener(addCordsAfterButtonPress(ptProduct = new Point(), productCords));
        btnSetGasketType.addActionListener(addCordsAfterButtonPress(ptGasketType = new Point(), gasketTypeCords));
        btnSeGasketThickness.addActionListener(addCordsAfterButtonPress(ptGasketThickness = new Point(), gasketThicknessCords));
        btnSetNps.addActionListener(addCordsAfterButtonPress(ptNps = new Point(), npsCords));
        btnSetAnsi.addActionListener(addCordsAfterButtonPress(ptAnsi = new Point(), ansiCords));
        btnSetGasketOd.addActionListener(addCordsAfterButtonPress(ptGasketOd = new Point(), gasketOdCords));
        btnSetGasketId.addActionListener(addCordsAfterButtonPress(ptGasketId = new Point(), gasketIdCords));
        btnSetBoltDiameter.addActionListener(addCordsAfterButtonPress(ptBoltHoleDia = new Point(), boltHoleDiaCords));
        btnSetBoltHoles.addActionListener(addCordsAfterButtonPress(ptNumOfBoltHoles = new Point(), numBoltHolesCords));
        btnSetWeight.addActionListener(addCordsAfterButtonPress(ptWeight = new Point(), weightCords));
        btnSetSku.addActionListener(addCordsAfterButtonPress(ptSku = new Point(), skuCords));
        btnSetPrice.addActionListener(addCordsAfterButtonPress(ptPrice = new Point(), priceCords));
        btnSetExcelTitleBar.addActionListener(addCordsAfterButtonPress(ptExcelTitle = new Point(), excelCords));
        btnSetOk.addActionListener(addCordsAfterButtonPress(ptOk = new Point(), okayCords));
        btnSetExcelEditCords.addActionListener(addCordsAfterButtonPress(ptExcelEdit = new Point(), excelEditCords));
        btnSetSave.addActionListener(addCordsAfterButtonPress(ptSave = new Point(), saveCords));

        btnStart.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    startSession();
                } catch (AWTException | InterruptedException | IOException | UnsupportedFlavorException awtException) {
                    awtException.printStackTrace();
                }
            }
        });

        // Show Window
        jFrame.setVisible(true);

    }

    private ActionListener addCordsAfterButtonPress(Point p, JLabel labelCords) {
        return new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                KeyListener keyListener = new KeyListener() {
                    @Override
                    public void keyTyped(KeyEvent e) {
                        p.setLocation(MouseInfo.getPointerInfo().getLocation());
                        labelCords.setText("[" + p.x + "," + p.y + "]");
                        jFrame.removeKeyListener(this);
                        jFrame.pack();
                    }

                    @Override
                    public void keyPressed(KeyEvent e) {

                    }

                    @Override
                    public void keyReleased(KeyEvent e) {

                    }
                };
                jFrame.addKeyListener(keyListener);
                jFrame.requestFocus();
            }
        };
    }

    public void startSession() throws AWTException, InterruptedException, IOException, UnsupportedFlavorException {

        // Start Robot
        Robot robot = new Robot();
        int i = 0;
        while ( i < Integer.parseInt(textFieldAmountOfEntry.getText())) {
            // Copy Excel Info
            // Click Excel
            leftClick(robot, ptExcelTitle);
            pressEsc(robot);
            pressDownArrow(robot);
            selectRow(robot);
            copy(robot);

            Clipboard c = Toolkit.getDefaultToolkit().getSystemClipboard();
            String data = (String) c.getData(DataFlavor.stringFlavor);
            String[] dataItems = data.split("\t");
            System.out.println(dataItems[0]);

            // Select Product

            leftClick(robot, ptVariation);
            leftClick(robot, ptVariation);
            selectDropDownMenuItem(robot, ptProduct, productOptions, dataItems[0]);
            selectDropDownMenuItem(robot, ptGasketType, gasketTypeOptions, dataItems[4]);
            selectDropDownMenuItem(robot, ptGasketThickness, gasketThicknessOptions, dataItems[5]);
            selectDropDownMenuItem(robot, ptNps, npsOptions, dataItems[6]);
            selectDropDownMenuItem(robot, ptAnsi, AnsiOptions, dataItems[7]);
            selectDropDownMenuItem(robot, ptGasketOd, gasketOdOptions, dataItems[8]);
            selectDropDownMenuItem(robot, ptGasketId, gacketIdOptions, dataItems[9]);
            selectDropDownMenuItem(robot, ptNumOfBoltHoles, numberOfBoltHolesOptions, dataItems[10]);
            selectDropDownMenuItem(robot, ptBoltHoleDia, boltHoleDiameterOptions, dataItems[11]);
            doubleClickAndPaste(robot, ptWeight, dataItems[14]);
            doubleClickAndPaste(robot, ptSku, dataItems[1]);
            doubleClickAndPaste(robot, ptPrice, dataItems[13].substring(1));
            leftClick(robot, ptOk);
            // While End
            i++;
        }

        leftClick(robot, ptSave);
        robot = null;
    }

    public void doubleClickAndPaste(Robot r, Point p, String paste) throws InterruptedException {

        StringSelection data = new StringSelection
                (paste);
        Clipboard cb = Toolkit.getDefaultToolkit()
                .getSystemClipboard();
        cb.setContents(data, data);

        r.mouseMove((int) p.getX(), (int) p.getY());
        Thread.sleep(1000);
        r.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        r.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(100);
        r.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        r.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1000);
        leftClick(r, ptExcelEdit);
        pressDownArrow(r);
        pressDownArrow(r);
        pressDownArrow(r);
        pressDownArrow(r);
        pressDownArrow(r);
        Thread.sleep(500);
        r.keyPress(KeyEvent.VK_ENTER);
        Thread.sleep(ONE_SECOND);
        r.keyRelease(KeyEvent.VK_ENTER);
    }

    public void selectDropDownMenuItem(Robot r, Point pos, String[] options, String optionSelected) throws InterruptedException {
        leftClick(r, pos);
        int i = 0;
        while (!optionSelected.equals(options[i])) {
            System.out.println(optionSelected + " - " + options[i]);
            pressDownArrow(r);
            i++;
        }

        Thread.sleep(250);
        r.keyPress(KeyEvent.VK_ENTER);
        Thread.sleep(ONE_SECOND);
        r.keyRelease(KeyEvent.VK_ENTER);
    }

    public void leftClick(Robot r, Point p) throws InterruptedException {
        r.mouseMove((int) p.getX(), (int) p.getY());
        Thread.sleep(ONE_SECOND);
        r.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        r.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void selectRow(Robot r) throws InterruptedException {
        Thread.sleep(ONE_SECOND);
        r.keyPress(KeyEvent.VK_SHIFT);
        Thread.sleep(ONE_SECOND);
        r.keyPress(KeyEvent.VK_SPACE);
        Thread.sleep(ONE_SECOND);
        r.keyRelease(KeyEvent.VK_SPACE);
        Thread.sleep(250);
        r.keyRelease(KeyEvent.VK_SHIFT);
        Thread.sleep(ONE_SECOND);
    }

    public void copy(Robot r) throws InterruptedException {
        leftClick(r, ptExcelEdit);
        pressDownArrow(r);
        pressDownArrow(r);
        pressDownArrow(r);
        pressDownArrow(r);
        Thread.sleep(500);
        r.keyPress(KeyEvent.VK_ENTER);
        Thread.sleep(ONE_SECOND);
        r.keyRelease(KeyEvent.VK_ENTER);
    }

    public void pressDownArrow(Robot r) throws InterruptedException {
        Thread.sleep(250);
        r.keyPress(KeyEvent.VK_DOWN);
        Thread.sleep(100);
        r.keyRelease(KeyEvent.VK_DOWN);
    }

    public void pressEsc(Robot r) throws InterruptedException {
        Thread.sleep(250);
        r.keyPress(KeyEvent.VK_ESCAPE);
        Thread.sleep(1000);
        r.keyRelease(KeyEvent.VK_ESCAPE);
    }

}
