package com.douma.entity;

import java.util.List;

public class MingxiEntity {

    private int index;
    private List<Integer> otherIndex;
    private String riqi;
    private String benji;
    private String kehu;

    public List<Integer> getOtherIndex() {
        return otherIndex;
    }

    public void setOtherIndex(List<Integer> otherIndex) {
        this.otherIndex = otherIndex;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getRiqi() {
        return riqi;
    }

    public void setRiqi(String riqi) {
        this.riqi = riqi;
    }

    public String getBenji() {
        return benji;
    }

    public void setBenji(String benji) {
        this.benji = benji;
    }

    public String getKehu() {
        return kehu;
    }

    public void setKehu(String kehu) {
        this.kehu = kehu;
    }
}
