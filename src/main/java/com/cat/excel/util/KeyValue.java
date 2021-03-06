package com.cat.excel.util;

public class KeyValue<K, V> {
	
	private K key;
	private V value;
	
	public KeyValue(K k, V v) {
		this.key = k;
		this.value = v;
	}
	
	public K getKey() {
		return key;
	}
	public void setKey(K key) {
		this.key = key;
	}
	public V getValue() {
		return value;
	}
	public void setValue(V value) {
		this.value = value;
	}
}
