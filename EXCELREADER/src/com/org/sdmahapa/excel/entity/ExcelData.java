package com.org.sdmahapa.excel.entity;

import java.util.Date;

public class ExcelData {
	private String name;
	private String description;
	private double price;
	private Date date;
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public double getPrice() {
		return price;
	}
	public void setPrice(double price) {
		this.price = price;
	}
	public Date getDate() {
		return date;
	}
	public void setDate(Date date) {
		this.date = date;
	}
	
	@Override
	public String toString(){
		return name+" Expanse for "+description+" at "+price+" INR "+" at time "+date.toString();
	}
	
	@Override
	public int hashCode()
	{
		return (name+description+price+date).hashCode();
	}
	
	@Override
	public boolean equals(Object obj){
		if (obj != null)
		{
			if (obj instanceof ExcelData)
			{
				ExcelData ed = (ExcelData)obj;
				if (ed.getName().equalsIgnoreCase(name) && ed.getPrice()==price)
					return true;
				else
					return false;
			}
			else
				return false;
		}
		return false;
	}
}
