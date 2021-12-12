package com.github.assemblathe1.production;

import lombok.Data;

@Data
public class ObjectToInsert {
    Double number;
    String organisation;
    String address;

    public ObjectToInsert(Double number, String organisation, String address) {
        this.number = number;
        this.organisation = organisation;
        this.address = address;
    }

    public Integer getNumber() {
        return number.intValue();
    }
}
