/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.mavenproject1.ragaiproject;

/**
 *
 * @author helios
 */
public class Pair {
    private String label;
    private String value;
    
    public Pair(String label, String value){
        this.label = label;
        this.value = value;
    }
    public String getLabel(){
        return this.label;
    }
    public String getValue(){
        return this.value;
    }
}
