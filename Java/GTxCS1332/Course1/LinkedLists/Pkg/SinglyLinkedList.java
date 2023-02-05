package Pkg;
import java.util.NoSuchElementException;

/**
 * Your implementation of a Singly-Linked List.
 */
public class SinglyLinkedList<T> {

    /*
     * Do not add new instance variables or modify existing ones.
     */
    private SinglyLinkedListNode<T> head;
    private SinglyLinkedListNode<T> tail;
    private int size;

    /*
     * Do not add a constructor.
     */

    /**
     * Adds the element to the front of the list.
     *
     * Method should run in O(1) time.
     *
     * @param data the data to add to the front of the list
     * @throws java.lang.IllegalArgumentException if data is null
     */
    public void addToFront(T data) {
        if (data == null) {
            throw new java.lang.IllegalArgumentException(); 
        } else {
            SinglyLinkedListNode<T> newNode = new SinglyLinkedListNode<T>(data);
            if (head == null) {
                tail = newNode;
            }
            else { 
                newNode.setNext(head); 
            }
            size++;
            head = newNode;
        }
    }

    public String toString() {
        String myString = "";
        SinglyLinkedListNode<T> n = head;
        Integer i = 1;
        if (head == null) { return "the list is empty!"; }
        if (head != null) { myString = myString + "H:" + (String) head.data + "\n"; }
        if (tail != null) { myString = myString + "T:" + (String) tail.data + "\n"; }
        while (n != null) {
            myString = myString + i.toString() + ":" + n.data + "\n";
            n = n.next;
            i++;
        }
        return myString;
    }

    /**
     * Adds the element to the back of the list.
     *
     * Method should run in O(1) time.
     *
     * @param data the data to add to the back of the list
     * @throws java.lang.IllegalArgumentException if data is null
     */
    public void addToBack(T data) {
        if (data == null) {
            throw new java.lang.IllegalArgumentException(); 
        } else {
            SinglyLinkedListNode<T> newNode = new SinglyLinkedListNode<T>(data);
            if (tail == null) {
                head = newNode;
            }
            else { 
                tail.setNext(newNode); 
            }
            size++;
            tail = newNode;
        }
    }

    /**
     * Removes and returns the first data of the list.
     *
     * Method should run in O(1) time.
     *
     * @return the data formerly located at the front of the list
     * @throws java.util.NoSuchElementException if the list is empty
     */
    public T removeFromFront() {
        if (head == null) {
            throw new java.util.NoSuchElementException();
        } else {
            T data = head.getData();
            if (head.getNext() == null) {
                head = null;
                tail = null;
            }
            else {
                head = head.getNext();
            }
            size--;
            return data;
        }
    }

    /**
     * Removes and returns the last data of the list.
     *
     * Method should run in O(n) time.
     *
     * @return the data formerly located at the back of the list
     * @throws java.util.NoSuchElementException if the list is empty
     */
    public T removeFromBack() {
        if (tail == null) {
            throw new java.util.NoSuchElementException();
        } else {
            T data = tail.getData();
            if (head == tail) {
                head = tail = null;
            } else {
                SinglyLinkedListNode<T> curr = head;
                while (curr.getNext() != tail) { //get second-to-last node
                    curr = curr.getNext();
                }
                curr.setNext(null);
                tail = curr;
            }
            size--;
            return data;
        }
    }
    /**
     * Returns the head node of the list.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return the node at the head of the list
     */
    public SinglyLinkedListNode<T> getHead() {
        // DO NOT MODIFY THIS METHOD!
        return head;
    }

    /**
     * Returns the tail node of the list.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return the node at the tail of the list
     */
    public SinglyLinkedListNode<T> getTail() {
        // DO NOT MODIFY THIS METHOD!
        return tail;
    }

    /**
     * Returns the size of the list.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return the size of the list
     */
    public int size() {
        // DO NOT MODIFY THIS METHOD!
        return size;
    }

    private class SinglyLinkedListNode<U> {

        private U data;
        private SinglyLinkedListNode<U> next;
    
        /**
         * Constructs a new SinglyLinkedListNode with the given data and next node
         * reference.
         *
         * @param data the data stored in the new node
         * @param next the next node in the list
         */
        SinglyLinkedListNode(U data, SinglyLinkedListNode<U> next) {
            this.data = data;
            this.next = next;
        }
    
        /**
         * Creates a new SinglyLinkedListNode with only the given data.
         *
         * @param data the data stored in the new node
         */
        SinglyLinkedListNode(U data) {
            this(data, null);
        }
    
        /**
         * Gets the data.
         *
         * @return the data
         */
        U getData() {
            return data;
        }
    
        /**
         * Gets the next node.
         *
         * @return the next node
         */
        SinglyLinkedListNode<U> getNext() {
            return next;
        }
    
        /**
         * Sets the next node.
         *
         * @param next the new next node
         */
        void setNext(SinglyLinkedListNode<U> next) {
            this.next = next;
        }
    }
}