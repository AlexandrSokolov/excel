package com.savdev.commons.excel.api;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.LocalTime;

public class TestDto {
  String colA;
  BigDecimal money;
  LocalDateTime someDate;
  LocalTime time;
  BigDecimal number;

  public String getColA() {
    return colA;
  }

  public void setColA(String colA) {
    this.colA = colA;
  }

  public BigDecimal getMoney() {
    return money;
  }

  public void setMoney(BigDecimal money) {
    this.money = money;
  }

  public LocalDateTime getSomeDate() {
    return someDate;
  }

  public void setSomeDate(LocalDateTime someDate) {
    this.someDate = someDate;
  }

  public LocalTime getTime() {
    return time;
  }

  public void setTime(LocalTime time) {
    this.time = time;
  }

  public BigDecimal getNumber() {
    return number;
  }

  public void setNumber(BigDecimal number) {
    this.number = number;
  }
}
