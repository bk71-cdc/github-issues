package cdc.gov.github.dto;

public class LoadTicketsRequest {
    private String fileName;
    
    public LoadTicketsRequest() {}
    
    public LoadTicketsRequest(String fileName) {
        this.fileName = fileName;
    }
    
    // Getters and Setters
    public String getFileName() {
        return fileName;
    }
    
    public void setFileName(String fileName) {
        this.fileName = fileName;
    }
    
    public boolean isValid() {
        return fileName != null && !fileName.trim().isEmpty();
    }
    
    @Override
    public String toString() {
        return "LoadTicketsRequest{" +
                "fileName='" + fileName + '\'' +
                '}';
    }
}
