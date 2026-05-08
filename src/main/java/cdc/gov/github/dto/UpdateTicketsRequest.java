package cdc.gov.github.dto;

public class UpdateTicketsRequest {
    private String fileName;
    
    public UpdateTicketsRequest() {}
    
    public UpdateTicketsRequest(String fileName) {
        this.fileName = fileName;
    }
    
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
        return "UpdateTicketsRequest{" +
                "fileName='" + fileName + '\'' +
                '}';
    }
}
