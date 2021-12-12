package com.github.assemblathe1.production;

import lombok.Data;

@Data
public class ObjectToInsert {
    int number;
    String organisation;
    String address;

    public ObjectToInsert(int number, String organisation, String address) {
        this.number = number;
        this.organisation = organisation;
        this.address = address;
    }

}
