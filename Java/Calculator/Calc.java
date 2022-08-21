import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.border.Border;
import java.util.Arrays;
import java.util.Scanner;

public class Calc implements ActionListener {

    private JFrame myFrame;
    private JLabel mainLabel;
    private JPanel mainPanel, buttonsPanel;
    private GridLayout myLayout;
    private String[] operators = {"+", "-", "*", "/"};
    private boolean isOutput;

    //URL iconURL = getClass().getResource("C:\\Users\\tusha\\Documents\\projects\\MyGithub\\tusharthomas\\Java\\Calculator\\R.pn")

    public Calc() {
        createCalcFrame();
        addCalcControls();
        myFrame.setVisible(true);
        isOutput = false;
    }

    private void createCalcFrame() {

        //frame specs, i.e. all constants
        final String TITLE = "Calculator";

        System.out.println("Height " + CalcDims.HEIGHT);
        System.out.println("Width " + CalcDims.WIDTH);

        myFrame = new JFrame();

        myLayout = new GridLayout(2, 1);
        myFrame.setLayout(myLayout);
        myFrame.setTitle(TITLE);
        myFrame.setBounds(CalcDims.LEFT, CalcDims.TOP, CalcDims.WIDTH, CalcDims.HEIGHT);
        myFrame.setResizable(false);
        myFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

    }

    private void addCalcControls() {

        mainPanel = createDefaultPanel();
        buttonsPanel = createDefaultPanel();

        //Main Panel: where output is shown
        mainLabel = new JLabel();
        mainLabel.setFont(new Font("Calibri", 0, 25));
        mainPanel.add(mainLabel);

        //Buttons
        buttonsPanel.setLayout(new GridLayout(5, 3));
        
        for (byte x = 1; x <= 9; x++) {addGenericButton("" + x);}
        addGenericButton("0");
        for (String s : operators) {addGenericButton(s);}
        addGenericButton("Go!");

    }

    private class ScreenDims {
        public static double height = Toolkit.getDefaultToolkit().getScreenSize().getSize().getHeight();
        public static double width = Toolkit.getDefaultToolkit().getScreenSize().getSize().getWidth();
    }

    private class CalcDims {
        final static int WIDTH = (int) (ScreenDims.width / 4);
        final static int HEIGHT = (int) (WIDTH * 1.33);
        final static int LEFT = (int) (ScreenDims.width/2 - WIDTH/2);
        final static int TOP = (int) (ScreenDims.height/2 - HEIGHT/2);
    }

    private JPanel createDefaultPanel() {

        Border paddingBorder = BorderFactory.createEmptyBorder(10, 10,10, 10);
        Border myLineBorder = BorderFactory.createLineBorder(new Color(0, 0,0 ));
        Border defaultBorder = BorderFactory.createCompoundBorder(paddingBorder, myLineBorder);

        JPanel myPanel = new JPanel();
        myPanel.setBorder(defaultBorder);

        myFrame.add(myPanel, BorderLayout.CENTER);

        return myPanel;
        
    }

    private void addGenericButton(String buttonCaption) {
        JButton myBtn = new JButton(buttonCaption);
        myBtn.addActionListener(this);
        buttonsPanel.add(myBtn);
    }

    private void updateLabel(String textToAppend) {

        boolean isOperator = Arrays.asList(operators).contains(textToAppend);

        if (isOutput) { mainLabel.setText(""); isOutput = false; }

        if (isOperator) {

            if (mainLabel.getText() == "")      {return;}
            
            if (getCurrentOperator() != "")     {mainLabel.setText(mainLabel.getText().replace(getCurrentOperator(), textToAppend));} 
            else                                {append(mainLabel, textToAppend + " ", " ");}

        } else {
            append(mainLabel, textToAppend, "");
        }

    }

    private String getCurrentOperator() {
        for (String s : operators) { if (mainLabel.getText().contains(s)) return s; }
        return "";
    }

    private void append(JLabel myLabel, String textToAppend, String delimiter) {
        myLabel.setText(myLabel.getText() + delimiter + textToAppend);
    }

    private void evaluateExpression() {

        if (getCurrentOperator() == "") {return;}

        Scanner myScanner = new Scanner(mainLabel.getText());
        String operator = getCurrentOperator().trim();
        int num1 = myScanner.nextInt();

        if (mainLabel.getText().replace(operator, "").replace(num1 + "", "").trim() == "") {return;}

        int num2 = Integer.valueOf(mainLabel.getText().replace(operator, "").replace(num1 + "", "").trim());
        float result;

        switch(operator) {
            case "+":       result = num1 + num2; break;
            case "-":       result = num1 - num2; break;
            case "*":       result = num1 * num2; break;
            case "/":       result = (float) num1 / (float) num2; break;
            default:        return;
        }

        mainLabel.setText(String.format("%.2f", result));

        isOutput = true;

    }

    @Override
    public void actionPerformed(ActionEvent e) {
        String buttonCaption = e.getActionCommand();
        if (buttonCaption == "Go!")     {evaluateExpression();}
        else                            {updateLabel(buttonCaption);}
    }

    public static void main(String[] args) {
        new Calc();
    }

}