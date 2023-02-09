package Pkg;

import java.util.List;
import java.util.ArrayList;

/**
 * Your implementation of the pre-order, in-order, and post-order
 * traversals of a tree.
 */
public class Traversals<T extends Comparable<? super T>> {

    /**
     * DO NOT ADD ANY GLOBAL VARIABLES!
     */

    /**
     * Given the root of a binary search tree, generate a
     * pre-order traversal of the tree. The original tree
     * should not be modified in any way.
     *
     * This must be done recursively.
     *
     * Must be O(n).
     *
     * @param <T> Generic type.
     * @param root The root of a BST.
     * @return List containing the pre-order traversal of the tree.
     */
    public List<T> preorder(TreeNode<T> root) {
        List<T> myList = new ArrayList<>();
        preorderRecurse(root, myList);
        return myList;
    }

    private void preorderRecurse(TreeNode<T> node, List<T> list) {
        if (node == null) {return;}
        list.add(node.getData());
        preorderRecurse(node.getLeft(), list);
        preorderRecurse(node.getRight(), list);
    }

    /**
     * Given the root of a binary search tree, generate an
     * in-order traversal of the tree. The original tree
     * should not be modified in any way.
     *
     * This must be done recursively.
     *
     * Must be O(n).
     *
     * @param <T> Generic type.
     * @param root The root of a BST.
     * @return List containing the in-order traversal of the tree.
     */
    public List<T> inorder(TreeNode<T> root) {
        List<T> myList = new ArrayList<>();
        inorderRecurse(root, myList);
        return myList;
    }

    private void inorderRecurse(TreeNode<T> node, List<T> list) {
        if (node == null) {return;}
        inorderRecurse(node.getLeft(), list);
        list.add(node.getData());
        inorderRecurse(node.getRight(), list);
    }

    /**
     * Given the root of a binary search tree, generate a
     * post-order traversal of the tree. The original tree
     * should not be modified in any way.
     *
     * This must be done recursively.
     *
     * Must be O(n).
     *
     * @param <T> Generic type.
     * @param root The root of a BST.
     * @return List containing the post-order traversal of the tree.
     */
    public List<T> postorder(TreeNode<T> root) {
        List<T> myList = new ArrayList<>();
        postorderRecurse(root, myList);
        return myList;
    }

    private void postorderRecurse(TreeNode<T> node, List<T> list) {
        if (node == null) {return;}
        postorderRecurse(node.getLeft(), list);
        postorderRecurse(node.getRight(), list);
        list.add(node.getData());
    }
}