import Pkg.TreeNode;
import Pkg.Traversals;
import java.util.List;

public class Main {
    
    public static void main(String[] args) {

        TreeNode<Integer> root = new TreeNode<>(10);
        Traversals<Integer> t = new Traversals<>();

        root.setLeft(new TreeNode<Integer>(5));
        root.setRight(new TreeNode<Integer>(15));

        List<Integer> list = t.inorder(root);

        System.out.println(list.toString());

    }

}
