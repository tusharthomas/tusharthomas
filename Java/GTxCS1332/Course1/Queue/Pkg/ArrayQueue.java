package Pkg;
import java.rmi.server.ObjID;
import java.util.NoSuchElementException;

/**
 * Your implementation of an ArrayQueue.
 */
public class ArrayQueue<T> {

    /*
     * The initial capacity of the ArrayQueue.
     *
     * DO NOT MODIFY THIS VARIABLE.
     */
    public static final int INITIAL_CAPACITY = 9;

    /*
     * Do not add new instance variables or modify existing ones.
     */
    private T[] backingArray;
    private int front;
    private int size;

    /**
     * This is the constructor that constructs a new ArrayQueue.
     * 
     * Recall that Java does not allow for regular generic array creation,
     * so instead we cast an Object[] to a T[] to get the generic typing.
     */
    public ArrayQueue() {
        // DO NOT MODIFY THIS METHOD!
        backingArray = (T[]) new Object[INITIAL_CAPACITY];
    }

    /**
     * Adds the data to the back of the queue.
     *
     * If sufficient space is not available in the backing array, resize it to
     * double the current length. When resizing, copy elements to the
     * beginning of the new array and reset front to 0.
     *
     * Method should run in amortized O(1) time.
     *
     * @param data the data to add to the back of the queue
     * @throws java.lang.IllegalArgumentException if data is null
     */
    public void enqueue(T data) {

        if (data == null) {
            throw new java.lang.IllegalArgumentException();
        }

        int addIndex = calculateAddIndex();

        if (backingArray[addIndex] != null) { 
            resize();
            addIndex = calculateAddIndex();
        }
        backingArray[addIndex] = data;
        size++;

    }

    private int calculateAddIndex() {
        return (front + size) % backingArray.length;
    }

    public String toString() {
        int i;
        String str = "";
        if (size == 0) {
            return "The queue is empty!";
        }
        for (i = 0; i < backingArray.length; i++) {
            str = str + backingArray[i];
            str = str + "\n";
        }
        return str;
    }

    private void resize() {
        T[] newArray = (T[]) new Object[backingArray.length * 2];
        int i, curr;
        for (i = 0; i < backingArray.length; i++) {
            curr = (front + i) % backingArray.length;
            newArray[i] = backingArray[curr];
        }
        front = 0;
        backingArray = newArray;
    }

    /**
     * Removes and returns the data from the front of the queue.
     *
     * Do not shrink the backing array.
     *
     * Replace any spots that you dequeue from with null.
     *
     * If the queue becomes empty as a result of this call, do not reset
     * front to 0.
     *
     * Method should run in O(1) time.
     *
     * @return the data formerly located at the front of the queue
     * @throws java.util.NoSuchElementException if the queue is empty
     */
    public T dequeue() {
        
        if (backingArray[front] == null) {
            throw new java.util.NoSuchElementException();
        }

        T data = backingArray[front];
        backingArray[front] = null;
        front = (front + 1 % backingArray.length);
        size--;

        return data;
        
    }

    /**
     * Returns the backing array of the queue.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return the backing array of the queue
     */
    public T[] getBackingArray() {
        // DO NOT MODIFY THIS METHOD!
        return backingArray;
    }

    /**
     * Returns the size of the queue.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return the size of the queue
     */
    public int size() {
        // DO NOT MODIFY THIS METHOD!
        return size;
    }
}