import MyPackage.ArrayList;

public class Main {

    public static void main(String[] args) {
        
        ArrayList<String> A = new ArrayList<>();
        
        A.addToFront("a1");
        A.addToFront("a2");
        A.addToFront("a1");
        A.addToFront("a2");
        A.addToFront("a1");
        A.addToFront("a2");
        A.addToFront("a1");
        A.addToFront("a2");
        A.addToFront("a1");

        A.removeFromFront();
        A.printArrayList();

    }

}
