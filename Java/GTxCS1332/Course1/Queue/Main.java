import Pkg.ArrayQueue;

public class Main {
    
    public static void main(String[] args) {

        ArrayQueue<String> Q = new ArrayQueue<>();

        Q.enqueue("A");
        Q.enqueue("B");
        Q.enqueue("C");

        System.out.println(Q.toString());

    }

}
