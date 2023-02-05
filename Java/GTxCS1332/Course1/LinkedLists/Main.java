import Pkg.SinglyLinkedList;

public class Main {

    public static void main(String[] args) {
    
        SinglyLinkedList<String> myList = new SinglyLinkedList<>();

        myList.addToFront("a");
        myList.addToFront("b");
        myList.addToBack("c");

        myList.removeFromFront();
        myList.removeFromFront();
        myList.removeFromFront();
        myList.removeFromFront();

        System.out.println(myList.toString());

    }
    
}
