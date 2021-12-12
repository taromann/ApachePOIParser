package com.github.assemblathe1.production;

import lombok.Data;

@Data
public class ObjectToInsert {
    String number;
    String organisation;
    String address;

    public ObjectToInsert(String number, String organisation, String address) {
        this.number = number;
        this.organisation = organisation;
        this.address = address;
    }

}
