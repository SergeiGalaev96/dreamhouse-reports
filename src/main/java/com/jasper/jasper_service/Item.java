package com.jasper.jasper_service;

public class Item {

    private String name;
    private int qty;
    private int price;

    public Item(String name, int qty, int price) {
        this.name = name;
        this.qty = qty;
        this.price = price;
    }

    public String getName() { return name; }
    public int getQty() { return qty; }
    public int getPrice() { return price; }
}