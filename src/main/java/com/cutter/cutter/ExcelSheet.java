package com.cutter.cutter;

import java.util.ArrayList;

public class ExcelSheet {
    private String filePath;
    private String fileName;
    private Integer totalCompanies;
    private String type;

    ArrayList<ArrayList<String>> terminationCodes;
    private ArrayList<String> headings;
    private ArrayList<Company> emailCompanies;
    private ArrayList<Company> cenCompanies;
    private ArrayList<Company> estCompanies;
    private ArrayList<Company> pacCompanies;

    public ExcelSheet(String filePath, String fileName, String type, ArrayList<String> headings, ArrayList<Company> emailCompanies, ArrayList<Company> cenCompanies, ArrayList<Company> estCompanies, ArrayList<Company> pacCompanies, ArrayList<ArrayList<String>> terminationCodes) {
        this.filePath = filePath;
        this.fileName = fileName;
        this.type = type;
        this.headings = headings;
        this.emailCompanies = emailCompanies;
        this.cenCompanies = cenCompanies;
        this.estCompanies = estCompanies;
        this.pacCompanies = pacCompanies;
        this.totalCompanies = this.cenCompanies.size() + this.estCompanies.size() + this.pacCompanies.size();
        this.terminationCodes = terminationCodes;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }
    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public ArrayList<String> getHeadings() {
        return headings;
    }

    public void setHeadings(ArrayList<String> headings) {
        this.headings = headings;
    }

    public ArrayList<Company> getEmailCompanies() {
        return emailCompanies;
    }

    public void setEmailCompanies(ArrayList<Company> emailCompanies) {
        this.emailCompanies = emailCompanies;
    }

    public Integer getTotalCompanies() {
        return totalCompanies;
    }

    public void setTotalCompanies(Integer totalCompanies) {
        this.totalCompanies = totalCompanies;
    }

    public ArrayList<Company> getCenCompanies() {
        return cenCompanies;
    }

    public void setCenCompanies(ArrayList<Company> cenCompanies) {
        this.cenCompanies = cenCompanies;
    }

    public ArrayList<Company> getEstCompanies() {
        return estCompanies;
    }

    public void setEstCompanies(ArrayList<Company> estCompanies) {
        this.estCompanies = estCompanies;
    }

    public ArrayList<Company> getPacCompanies() {
        return pacCompanies;
    }

    public void setPacCompanies(ArrayList<Company> pacCompanies) {
        this.pacCompanies = pacCompanies;
    }

    public ArrayList<ArrayList<String>> getTerminationCodes() {
        return terminationCodes;
    }

    public void setTerminationCodes(ArrayList<ArrayList<String>> terminationCodes) {
        this.terminationCodes = terminationCodes;
    }
}
