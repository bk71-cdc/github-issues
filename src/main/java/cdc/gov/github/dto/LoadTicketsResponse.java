package cdc.gov.github.dto;

import java.util.ArrayList;
import java.util.List;

public class LoadTicketsResponse {
    private int totalCount;
    private int successCount;
    private int failureCount;
    private String message;
    private List<String> errors;
    private List<String> issueIds;
    
    public LoadTicketsResponse() {
        this.errors = new ArrayList<>();
        this.issueIds = new ArrayList<>();
    }
    
    public LoadTicketsResponse(int totalCount, int successCount, int failureCount, String message) {
        this.totalCount = totalCount;
        this.successCount = successCount;
        this.failureCount = failureCount;
        this.message = message;
        this.errors = new ArrayList<>();
        this.issueIds = new ArrayList<>();
    }
    
    // Getters and Setters
    public int getTotalCount() {
        return totalCount;
    }
    
    public void setTotalCount(int totalCount) {
        this.totalCount = totalCount;
    }
    
    public int getSuccessCount() {
        return successCount;
    }
    
    public void setSuccessCount(int successCount) {
        this.successCount = successCount;
    }
    
    public int getFailureCount() {
        return failureCount;
    }
    
    public void setFailureCount(int failureCount) {
        this.failureCount = failureCount;
    }
    
    public String getMessage() {
        return message;
    }
    
    public void setMessage(String message) {
        this.message = message;
    }
    
    public List<String> getErrors() {
        return errors;
    }
    
    public void setErrors(List<String> errors) {
        this.errors = errors;
    }
    
    public List<String> getIssueIds() {
        return issueIds;
    }
    
    public void setIssueIds(List<String> issueIds) {
        this.issueIds = issueIds;
    }
    
    public void addError(String error) {
        this.errors.add(error);
    }
    
    public void addIssueId(String issueId) {
        this.issueIds.add(issueId);
    }
    
    @Override
    public String toString() {
        return "LoadTicketsResponse{" +
                "totalCount=" + totalCount +
                ", successCount=" + successCount +
                ", failureCount=" + failureCount +
                ", message='" + message + '\'' +
                ", errors=" + errors +
                ", issueIds=" + issueIds +
                '}';
    }
}
