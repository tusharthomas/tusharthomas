package Pkg;

import java.util.ArrayList;
import java.util.List;

/**
 * Your implementation of a BST.
 */
public class BST<T extends Comparable<? super T>> {

    /*
     * Do not add new instance variables or modify existing ones.
     */
    private BSTNode<T> root;
    private int size;

    /*
     * Do not add a constructor.
     */

    /**
     * Adds the data to the tree.
     *
     * This must be done recursively.
     *
     * The new data should become a leaf in the tree.
     *
     * Traverse the tree to find the appropriate location. If the data is
     * already in the tree, then nothing should be done (the duplicate
     * shouldn't get added, and size should not be incremented).
     *
     * Should be O(log n) for best and average cases and O(n) for worst case.
     *
     * @param data The data to add to the tree.
     * @throws java.lang.IllegalArgumentException If data is null.
     */
    public void add(T data) {
        if (data == null) {
            throw new java.lang.IllegalArgumentException();
        } else {
            root = rAdd(data, root);
        }
        
    }

    private BSTNode<T> rAdd(T data, BSTNode<T> node) {
        if (node == null) {
            size++;
            return new BSTNode<T>(data);
        } else {
            T nodeData = node.getData();
            if (data.equals(nodeData)) {
                return node;
            }
            else if (data.compareTo(nodeData) > 0) {                 //data is greater
                node.setRight(rAdd(data, node.getRight()));
            } else {                                            //data is lesser
                node.setLeft(rAdd(data, node.getLeft()));
            }
            return node;
        }
    }

    public List<T> inorder() {
        List<T> myList = new ArrayList<>();
        inorderRecurse(root, myList);
        return myList;
    }

    private void inorderRecurse(BSTNode<T> node, List<T> list) {
        if (node == null) {return;}
        inorderRecurse(node.getLeft(), list);
        list.add(node.getData());
        inorderRecurse(node.getRight(), list);
    }

    /**
     * Removes and returns the data from the tree matching the given parameter.
     *
     * This must be done recursively.
     *
     * There are 3 cases to consider:
     * 1: The node containing the data is a leaf (no children). In this case,
     * simply remove it.
     * 2: The node containing the data has one child. In this case, simply
     * replace it with its child.
     * 3: The node containing the data has 2 children. Use the SUCCESSOR to
     * replace the data. You should use recursion to find and remove the
     * successor (you will likely need an additional helper method to
     * handle this case efficiently).
     *
     * Do NOT return the same data that was passed in. Return the data that
     * was stored in the tree.
     *
     * Hint: Should you use value equality or reference equality?
     *
     * Must be O(log n) for best and average cases and O(n) for worst case.
     *
     * @param data The data to remove.
     * @return The data that was removed.
     * @throws java.lang.IllegalArgumentException If data is null.
     * @throws java.util.NoSuchElementException   If the data is not in the tree.
     */
    public void remove(T data) {

        if (data == null) {
            throw new java.lang.IllegalArgumentException();
        }

        root = rRemove(data, root);

    }

    private BSTNode<T> rRemove(T data, BSTNode<T> node) {

        if (node == null) {
            throw new java.util.NoSuchElementException();
        }

        T nodeData = node.getData();
        if (data.equals(nodeData)) {    //found node to remove

            int countChildren = numChildren(node);

            if (countChildren == 0) {
                size--;
                return null;
            } else if (countChildren == 1) {
                size--;
                return getOnlyChild(node);
            } else {
                T successorData = getSuccessorData(node);
                node = rRemove(successorData, node);
                node.setData(successorData);
            }

        } else if (data.compareTo(nodeData) > 0) {  //data > node data
            node.setRight(rRemove(data, node.getRight()));
        } else {    //data < node data
            node.setLeft(rRemove(data, node.getLeft()));
        }

        return node;

    }

    private int numChildren(BSTNode<T> node) {
        int count = 0;
        if (node.getLeft() != null) {count++;}
        if (node.getRight() != null) {count++;}
        return count;
    }

    private BSTNode<T> getOnlyChild(BSTNode<T> parent) {
        if (parent.getLeft() != null) {
            return parent.getLeft();
        } else {
            return parent.getRight();
        }
    }

    private T getSuccessorData(BSTNode<T> node) {
        BSTNode<T> successor = node.getRight();
        while (successor.getLeft() != null) {
            successor = successor.getLeft();
        }
        return successor.getData();
    }

    public String toString() {
        List<T> data = inorder();
        return data.toString();
    }

    /**
     * Returns the root of the tree.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return The root of the tree
     */
    public BSTNode<T> getRoot() {

        // DO NOT MODIFY THIS METHOD!
        return root;
    }

    /**
     * Returns the size of the tree.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return The size of the tree
     */
    public int size() {
        // DO NOT MODIFY THIS METHOD!
        return size;
    }
}