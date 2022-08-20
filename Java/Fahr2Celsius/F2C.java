import java.util.Scanner;

public class F2C {
    public static void main(String[] args) {
        int f = getInputInt("Enter a Fahrenheit value: ");
        double c = (5.0/9) * (f - 32);
        System.out.printf("%-25s %.1f\n","Celsius:", c);
    }

    public static int getInputInt(String prompt) {
        
        Scanner input = new Scanner(System.in);
        int i;

        System.out.print(prompt);

        try {
            i = input.nextInt();
        } catch (Exception e) {
            System.out.println("Please enter a number!");
            i = getInputInt(prompt);
        }

        input.close();

        return i;

    }
}
