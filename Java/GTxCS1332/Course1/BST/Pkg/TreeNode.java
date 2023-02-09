package Pkg;
/**
 * Node class used for implementing the BST.
 *
 * DO NOT MODIFY THIS FILE!!
 *
 * @author CS 1332 TAs
 * @version 1.0
 */
public class TreeNode<T extends Comparable<? super T>> {

    private T data;
    private TreeNode<T> left;
    private TreeNode<T> right;

    /**
     * Constructs a TreeNode with the given data.
     *
     * @param data the data stored in the new node
     */
    public TreeNode(T data) {
        this.data = data;
    }

    /**
     * Gets the data.
     *
     * @return the data
     */
    public T getData() {
        return data;
    }

    /**
     * Gets the left child.
     *
     * @return the left child
     */
    public TreeNode<T> getLeft() {
        return left;
    }

    /**
     * Gets the right child.
     *
     * @return the right child
     */
    public TreeNode<T> getRight() {
        return right;
    }

    /**
     * Sets the data.
     *
     * @param data the new data
     */
    public void setData(T data) {
        this.data = data;
    }

    /**
     * Sets the left child.
     *
     * @param left the new left child
     */
    public void setLeft(TreeNode<T> left) {
        this.left = left;
    }

    /**
     * Sets the right child.
     *
     * @param right the new right child
     */
    public void setRight(TreeNode<T> right) {
        this.right = right;
    }
}