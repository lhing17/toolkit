import java.util.Iterator;

/**
 * 基于链表实现的栈
 * @author G.Seinfeld
 * @date 2018/6/15
 */
public class NodeStack<T> implements Iterable<T> {
    private Node first;
    private int size;

    @Override
    public Iterator<T> iterator() {
        return new Iterator<T>() {

            private Node current = first;

            @Override
            public boolean hasNext() {
                return current != null;
            }

            @Override
            public T next() {
                T item = current.item;
                current = current.next;
                return item;
            }
        };
    }

    public boolean isEmpty() {
        return size == 0;
    }

    public int size() {
        return size;
    }

    public void push(T item) {
        Node oldFirst = first;
        first = new Node();
        first.item = item;
        first.next = oldFirst;
        size++;
    }

    public T pop() {
        T item = first.item;
        first = first.next;
        size--;
        return item;
    }

    public T peek() {
        return first.item;
    }

    private class Node {
        private T item;
        private Node next;
    }
}
