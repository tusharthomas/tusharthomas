import Pkg.BST;

public class Main {
    
    public static void main(String[] args) {

        BST<Integer> myTree = new BST<>();

        myTree.add(1);
        myTree.add(0);
        myTree.add(2);
        myTree.add(0);

        System.out.println(myTree.toString());

    }

}
