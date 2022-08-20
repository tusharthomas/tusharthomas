import java.util.Scanner;

public class F2C {
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.print("Enter a Fahrenheit value: ");
        int f = input.nextInt();
        input.close();
        double c = (5.0/9) * (f - 32);
        System.out.println("Celsius: " + c);
    }
}
