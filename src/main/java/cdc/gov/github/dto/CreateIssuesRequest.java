package cdc.gov.github.dto;

public class CreateIssuesRequest {
    private int numberOfIssues;
    
    public CreateIssuesRequest() {}
    
    public CreateIssuesRequest(int numberOfIssues) {
        this.numberOfIssues = numberOfIssues;
    }
    
    public int getNumberOfIssues() {
        return numberOfIssues;
    }
    
    public void setNumberOfIssues(int numberOfIssues) {
        this.numberOfIssues = numberOfIssues;
    }
    
    public boolean isValid() {
        return numberOfIssues > 0;
    }
    
    @Override
    public String toString() {
        return "CreateIssuesRequest{" +
                "numberOfIssues=" + numberOfIssues +
                '}';
    }
}
