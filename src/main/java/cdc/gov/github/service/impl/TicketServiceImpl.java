package cdc.gov.github.service.impl;

import cdc.gov.github.dto.LoadTicketsResponse;
import cdc.gov.github.service.TicketService;
import com.google.gson.*;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.core.io.ClassPathResource;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class TicketServiceImpl implements TicketService {
    
    private static final String GITHUB_API_URL = "https://api.github.com/graphql";
    private static final String GITHUB_TOKEN = "";
    private static final String OWNER = "bk71-cdc";
    private static final String REPO = "github-issues";
    private static final String PROJECT_NAME = "@GitHub_Jira";
    
    private final OkHttpClient client = new OkHttpClient();
    private final Gson gson = new GsonBuilder().setPrettyPrinting().create();
    private String repositoryId = "";
    private String projectId = "";
    private String ownerId = "";

    @Override
    public LoadTicketsResponse loadTicketsFromExcel(String fileName) {
        LoadTicketsResponse response = new LoadTicketsResponse();
        
        try {
            // Initialize repository info
            initializeRepositoryInfo();
            
            // Get the Excel file from resources or use provided path
            InputStream excelStream = getExcelInputStream(fileName);
            if (excelStream == null) {
                response.setMessage("Excel file not found: " + fileName);
                response.addError("File not found: " + fileName);
                return response;
            }
            
            // Process the Excel file
            List<String> createdIssueIds = processExcelFile(excelStream, response);
            
            response.setTotalCount(createdIssueIds.size());
            response.setSuccessCount(createdIssueIds.size());
            response.setFailureCount(0);
            response.setIssueIds(createdIssueIds);
            response.setMessage("Successfully created " + createdIssueIds.size() + " issues from " + fileName);
            
        } catch (Exception e) {
            response.setMessage("Error loading tickets from Excel: " + e.getMessage());
            response.addError("Exception: " + e.getMessage());
            response.setTotalCount(0);
            response.setSuccessCount(0);
            response.setFailureCount(1);
        }
        
        return response;
    }

    @Override
    public LoadTicketsResponse createEmptyIssues(int numberOfIssues) {
        LoadTicketsResponse response = new LoadTicketsResponse();
        
        try {
            // Initialize repository info
            initializeRepositoryInfo();
            
            // Validate input
            if (numberOfIssues <= 0) {
                response.setMessage("Number of issues must be greater than 0");
                return response;
            }
            
            List<String> createdIssueIds = new ArrayList<>();
            
            // Create empty issues using improvement_form.yml structure
            for (int i = 0; i < numberOfIssues; i++) {
                try {
                    String issueTitle = "Empty Issue #" + (i + 1);
                    String issueBody = populateImprovementFormTemplate(
                        issueTitle, "eHARS", "Task", "Average", 
                        "Release 4.20", "Release 4.21", "All Document Import", "No Module", 
                        "", "This is an auto-generated empty issue.", "", "", "");
                    
                    // Call GitHub API to create issue
                    String[] issueResult = createIssue(repositoryId, issueTitle, issueBody, 
                        Arrays.asList("improvement"), new ArrayList<>(), "");
                    
                    if (issueResult != null && issueResult.length == 2) {
                        String issueId = issueResult[0];
                        String issueNumber = issueResult[1];
                        createdIssueIds.add(issueId);
                        // Add to project
                        addIssueToProject(projectId, issueId, "Todo", "Task", "Release 4.21", "", "", "eHARS");
                    } else {
                        response.addError("Failed to create issue #" + (i + 1));
                    }
                } catch (Exception e) {
                    response.addError("Error creating issue #" + (i + 1) + ": " + e.getMessage());
                }
            }
            
            response.setTotalCount(numberOfIssues);
            response.setSuccessCount(createdIssueIds.size());
            response.setFailureCount(numberOfIssues - createdIssueIds.size());
            response.setIssueIds(createdIssueIds);
            response.setMessage("Created " + createdIssueIds.size() + " empty issues out of " + numberOfIssues + " requested");
            
        } catch (Exception e) {
            response.setMessage("Error creating empty issues: " + e.getMessage());
            response.addError("Exception: " + e.getMessage());
            response.setTotalCount(numberOfIssues);
            response.setSuccessCount(0);
            response.setFailureCount(numberOfIssues);
        }
        
        return response;
    }

    @Override
    public LoadTicketsResponse updateTicketsFromExcel(String fileName) {
        LoadTicketsResponse response = new LoadTicketsResponse();
        
        try {
            // Initialize repository info
            initializeRepositoryInfo();
            
            // Get the Excel file from resources or use provided path
            InputStream excelStream = getExcelInputStream(fileName);
            if (excelStream == null) {
                response.setMessage("Excel file not found: " + fileName);
                response.addError("File not found: " + fileName);
                return response;
            }
            
            // Process the Excel file for updates
            List<String> updatedIssueIds = processExcelFileForUpdates(excelStream, response);
            
            response.setTotalCount(updatedIssueIds.size());
            response.setSuccessCount(updatedIssueIds.size());
            response.setFailureCount(0);
            response.setIssueIds(updatedIssueIds);
            response.setMessage("Successfully updated " + updatedIssueIds.size() + " tickets from " + fileName);
            
        } catch (Exception e) {
            response.setMessage("Error updating tickets from Excel: " + e.getMessage());
            response.addError("Exception: " + e.getMessage());
            response.setTotalCount(0);
            response.setSuccessCount(0);
            response.setFailureCount(1);
        }
        
        return response;
    }
    
    private void initializeRepositoryInfo() throws IOException {
        if (repositoryId.isEmpty()) {
            System.out.println("Initializing repository info...");
            System.out.println("Owner: " + OWNER);
            System.out.println("Repo: " + REPO);
            System.out.println("Project: " + PROJECT_NAME);
            System.out.println("Token: " + (GITHUB_TOKEN.length() > 10 ? GITHUB_TOKEN.substring(0, 10) + "..." : "INVALID"));
            
            try {
                JsonObject repoInfo = getRepoInfo(OWNER, REPO, PROJECT_NAME);
                System.out.println("Repo info response: " + repoInfo.toString());
                
                repositoryId = repoInfo.get("repoId").getAsString();
                projectId = repoInfo.get("projectId").getAsString();
                ownerId = repoInfo.get("ownerId").getAsString();
                
                System.out.println("Repository ID: " + repositoryId);
                System.out.println("Project ID: " + projectId);
                System.out.println("Owner ID: " + ownerId);
            } catch (Exception e) {
                System.err.println("Error initializing repository info: " + e.getMessage());
                throw e;
            }
        }
    }
    
    private InputStream getExcelInputStream(String fileName) {
        try {
            // First try to load from resources
            ClassPathResource resource = new ClassPathResource(fileName);
            if (resource.exists()) {
                System.out.println("Found file in resources: " + fileName);
                return resource.getInputStream();
            }
            
            // Try as absolute path
            File file = new File(fileName);
            if (file.exists()) {
                System.out.println("Found file as absolute path: " + file.getAbsolutePath());
                return new FileInputStream(file);
            }
            
            // Try as relative path from project root
            File relativeFile = new File(System.getProperty("user.dir"), fileName);
            if (relativeFile.exists()) {
                System.out.println("Found file as relative path: " + relativeFile.getAbsolutePath());
                return new FileInputStream(relativeFile);
            }
            
            System.out.println("File not found: " + fileName);
            System.out.println("Current working directory: " + System.getProperty("user.dir"));
            
            return null;
        } catch (Exception e) {
            System.err.println("Error loading file " + fileName + ": " + e.getMessage());
            return null;
        }
    }
    
    private List<String> processExcelFile(InputStream excelStream, LoadTicketsResponse response) throws IOException {
        List<String> createdIssueIds = new ArrayList<>();
        
        Workbook workbook;
        try {
            // Detect Excel format and use appropriate workbook
            if (excelStream.markSupported()) {
                excelStream.mark(0);
                // Try to detect if it's XLS or XLSX
                byte[] header = new byte[8];
                int bytesRead = excelStream.read(header);
                excelStream.reset();
                
                // Check for XLS signature (OLE2 format)
                if (bytesRead >= 8 && 
                    header[0] == 0xD0 && header[1] == 0xCF && 
                    header[2] == 0x11 && header[3] == 0xE0 && 
                    header[4] == 0xA1 && header[5] == 0xB1 && 
                    header[6] == 0x1A && header[7] == 0xE1) {
                    workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
                    System.out.println("Detected XLS format, using HSSFWorkbook");
                } else {
                    // Try XLSX first, if it fails, fall back to XLS
                    try {
                        workbook = new XSSFWorkbook(excelStream);
                        System.out.println("Detected XLSX format, using XSSFWorkbook");
                    } catch (Exception xlsxException) {
                        // Reset stream and try XLS
                        excelStream.reset();
                        workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
                        System.out.println("XLSX failed, falling back to XLS format with HSSFWorkbook");
                    }
                }
            } else {
                // If stream doesn't support mark, try XLSX first, then XLS
                try {
                    workbook = new XSSFWorkbook(excelStream);
                    System.out.println("Stream doesn't support mark, trying XSSFWorkbook first - succeeded");
                } catch (Exception xlsxException) {
                    // Create a fresh stream for XLS
                    workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
                    System.out.println("XSSFWorkbook failed, using HSSFWorkbook");
                }
            }
        } catch (Exception e) {
            throw new IOException("Failed to create workbook: " + e.getMessage(), e);
        }
        
        try {
            Sheet sheet = workbook.getSheetAt(0);
            
            // Find header row and capture ALL column names
            Row headerRow = sheet.getRow(3); // Header is on row 4 (0-indexed is 3)
            List<String> columnNames = new ArrayList<>();
            Map<String, Integer> columnMap = new HashMap<>();
            
            // Capture all column names and find key columns
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String columnName = cell.getStringCellValue().trim();
                    columnNames.add(columnName);
                    columnMap.put(columnName.toLowerCase(), i);
                }
            }
            
            System.out.println("Found columns: " + columnNames);
            System.out.println("Column map: " + columnMap);
            System.out.println("Total rows in sheet: " + sheet.getLastRowNum());
            
            // Process data rows, starting from row 5 (0-indexed is 4)
            for (int rowNum = 4; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;
                
                // Extract data from key columns
                String title = getCellValue(row, columnMap, "summary");
                System.out.println("Row " + rowNum + " - Title: '" + title + "'");
                if (title.isEmpty()) {
                    System.out.println("Skipping row " + rowNum + " - empty title");
                    continue;
                }
                
                String project = getCellValue(row, columnMap, "project");
                String issueType = getCellValue(row, columnMap, "issue type");
                String priority = getCellValue(row, columnMap, "priority");
                String affectsVersion = getCellValue(row, columnMap, "affects version/s");
                String fixVersion = getCellValue(row, columnMap, "fix version/s");
                String components = getCellValue(row, columnMap, "component/s");
                String moduleInfo = getCellValue(row, columnMap, "module info");
                String assignee = getCellValue(row, columnMap, "assignee");
                String reporter = getCellValue(row, columnMap, "reporter");
                String status = getCellValue(row, columnMap, "status");
                String description = getCellValue(row, columnMap, "description");
                String comments = getCellValue(row, columnMap, "comments");
                String linkedIssues = getCellValue(row, columnMap, "linked issues");
                String attachment = getCellValue(row, columnMap, "attachment");
                String resolution = getCellValue(row, columnMap, "resolution");
                
                // Map issue type to label
                String mappedIssueType = mapIssueTypeLabel(issueType);
                
                // Set default project since not in Excel headers
                String projectValue = project != null && !project.isEmpty() ? project : "eHARS";
                
                // Create issue body using improvement_form.yml structure
                String body = populateImprovementFormTemplate(
                    title, projectValue, mappedIssueType, priority, 
                    affectsVersion, fixVersion, components, moduleInfo, 
                    assignee, description, linkedIssues, attachment, resolution);
                
                // Create GitHub issue with proper labels based on issue type
                List<String> labels = new ArrayList<>();
                if (!mappedIssueType.isEmpty()) {
                    labels.add(mappedIssueType);
                } else {
                    labels.add("enhancement"); // default if issue type mapping fails
                }
                
                List<String> assignees = new ArrayList<>();
                if (!assignee.isEmpty()) {
                    String[] assigneeList = assignee.split(",");
                    for (String a : assigneeList) {
                        assignees.add(a.trim());
                    }
                }
                
                String[] issueResult = createIssue(repositoryId, title, body, labels, assignees, "");
                if (issueResult != null && issueResult.length == 2) {
                    String issueId = issueResult[0];
                    String issueNumber = issueResult[1];
                    
                    // Close issue if status is "closed"
                    if ("closed".equalsIgnoreCase(status)) {
                        try {
                            closeIssue(issueId);
                        } catch (Exception e) {
                            System.out.println("Warning: Failed to close issue #" + issueNumber + ": " + e.getMessage());
                        }
                    }
                    
                    createdIssueIds.add("Issue #" + issueNumber);
                    
                    // Add second issue body with comments as the very last step
                    if (comments != null && !comments.isEmpty()) {
                        try {
                            addSecondIssueBody(issueId, issueNumber, comments);
                        } catch (Exception e) {
                            System.out.println("Warning: Failed to add second issue body to issue #" + issueNumber + ": " + e.getMessage());
                        }
                    }
                } else {
                    response.addError("Failed to create issue for row " + rowNum + ": " + title);
                }
            }
            
            workbook.close();
        } catch (Exception e) {
            throw new IOException("Failed to process Excel file: " + e.getMessage(), e);
        }
        
        return createdIssueIds;
    }
    
    private List<String> processExcelFileForUpdates(InputStream excelStream, LoadTicketsResponse response) throws IOException {
        List<String> updatedIssueIds = new ArrayList<>();
        
        try (Workbook workbook = new XSSFWorkbook(excelStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            
            // Find header row and capture ALL column names
            Row headerRow = sheet.getRow(0);
            List<String> columnNames = new ArrayList<>();
            Map<String, Integer> columnMap = new HashMap<>();
            
            // Capture all column names and find key columns
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String columnName = cell.getStringCellValue().trim();
                    columnNames.add(columnName);
                    columnMap.put(columnName.toLowerCase(), i);
                }
            }
            
            // Process data rows
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;
                
                // Extract issue ID (key field like EHARS-600)
                String issueId = getCellValue(row, columnMap, "id");
                String ticketId = getCellValue(row, columnMap, "ticket id");
                String key = getCellValue(row, columnMap, "key");
                
                // Try to find issue number from any of these columns
                String issueNumber = extractIssueNumber(issueId);
                if (issueNumber.isEmpty()) {
                    issueNumber = extractIssueNumber(ticketId);
                }
                if (issueNumber.isEmpty()) {
                    issueNumber = extractIssueNumber(key);
                }
                
                if (issueNumber.isEmpty()) {
                    response.addError("No valid issue ID found in row " + rowNum + ". Expected format like EHARS-600");
                    continue;
                }
                
                // Extract other data from key columns
                String title = getCellValue(row, columnMap, "title");
                String project = getCellValue(row, columnMap, "project");
                String issueType = getCellValue(row, columnMap, "issue type");
                String priority = getCellValue(row, columnMap, "priority");
                String affectsVersion = getCellValue(row, columnMap, "affects version");
                String fixVersion = getCellValue(row, columnMap, "fix version");
                String components = getCellValue(row, columnMap, "components");
                String moduleInfo = getCellValue(row, columnMap, "module info");
                String assignee = getCellValue(row, columnMap, "assignee");
                String description = getCellValue(row, columnMap, "description");
                String comments = getCellValue(row, columnMap, "comments");
                String linkedIssues = getCellValue(row, columnMap, "linked issues");
                String attachment = getCellValue(row, columnMap, "attachment");
                String resolution = getCellValue(row, columnMap, "resolution");
                
                // Map issue type to label
                String mappedIssueType = mapIssueTypeLabel(issueType);
                
                // Create issue body using improvement_form.yml structure
                String body = populateImprovementFormTemplate(
                    title.isEmpty() ? "Updated Issue #" + issueNumber : title, 
                    project, mappedIssueType, priority, 
                    affectsVersion, fixVersion, components, moduleInfo, 
                    assignee, description, linkedIssues, attachment, resolution);
                
                // Update existing GitHub issue
                try {
                    String updatedIssueId = updateIssue(repositoryId, issueNumber, title, body, mappedIssueType, assignee);
                    if (updatedIssueId != null && !updatedIssueId.isEmpty()) {
                        updatedIssueIds.add(updatedIssueId);
                        
                        // Handle issue status - close if "closed", otherwise keep open
                        String status = getCellValue(row, columnMap, "status");
                        if ("closed".equalsIgnoreCase(status)) {
                            try {
                                closeIssue(updatedIssueId);
                            } catch (Exception e) {
                                System.out.println("Warning: Failed to close issue #" + issueNumber + ": " + e.getMessage());
                            }
                        }
                        
                        // Add second issue body with comments as the very last step
                        if (comments != null && !comments.isEmpty()) {
                            try {
                                addSecondIssueBody(updatedIssueId, issueNumber, comments);
                            } catch (Exception e) {
                                System.out.println("Warning: Failed to add second issue body to issue #" + issueNumber + ": " + e.getMessage());
                            }
                        }
                    } else {
                        response.addError("Failed to update issue #" + issueNumber + " for row " + rowNum);
                    }
                } catch (Exception e) {
                    response.addError("Error updating issue #" + issueNumber + " for row " + rowNum + ": " + e.getMessage());
                }
            }
            
            workbook.close();
        }
        
        return updatedIssueIds;
    }
    
    private String extractIssueNumber(String issueId) {
        if (issueId == null || issueId.isEmpty()) return "";
        
        // Pattern to match formats like "EHARS-600", "PROJ-123", etc.
        // Extract the number after the last hyphen
        int lastHyphen = issueId.lastIndexOf('-');
        if (lastHyphen > 0 && lastHyphen < issueId.length() - 1) {
            String numberPart = issueId.substring(lastHyphen + 1);
            try {
                // Validate it's a number
                Integer.parseInt(numberPart);
                return numberPart;
            } catch (NumberFormatException e) {
                return "";
            }
        }
        
        return "";
    }
    
    private String updateIssue(String repoId, String issueNumber, String title, String body, String issueType, String assignee) throws IOException {
        String mutation = """
            mutation($input: UpdateIssueInput!) {
              updateIssue(input: $input) {
                issue {
                  id
                  number
                }
              }
            }
            """;
        
        JsonObject input = new JsonObject();
        input.addProperty("repositoryId", repoId);
        input.addProperty("issueNumber", Integer.parseInt(issueNumber));
        
        if (title != null && !title.isEmpty()) {
            input.addProperty("title", title);
        }
        
        if (body != null && !body.isEmpty()) {
            input.addProperty("body", body);
        }
        
        JsonObject variables = new JsonObject();
        variables.add("input", input);
        
        JsonObject response = executeGraphQL(mutation, variables);
        
        // Debug: Print the full response
        System.out.println("GraphQL Response: " + response.toString());
        
        if (response == null) {
            System.out.println("Response is null");
            return null;
        }
        
        // Check for errors in GraphQL response
        if (response.has("errors")) {
            JsonArray errors = response.getAsJsonArray("errors");
            System.out.println("GraphQL Errors: " + errors.toString());
            return null;
        }
        
        JsonElement data = response.get("data");
        if (data == null || data.isJsonNull()) {
            System.out.println("No data field in response");
            return null;
        }
        
        if (data.isJsonObject() && data.getAsJsonObject().has("updateIssue")) {
            JsonElement updateIssueElement = new Gson().fromJson(data.getAsJsonObject().get("updateIssue").toString(), JsonElement.class);
            if (updateIssueElement.isJsonNull()) {
                System.out.println("updateIssue is null");
                return null;
            }
            
            JsonObject updateIssue = updateIssueElement.getAsJsonObject();
            if (updateIssue.has("issue")) {
                JsonElement issueElement = updateIssue.get("issue");
                if (issueElement.isJsonNull()) {
                    System.out.println("issue is null in updateIssue");
                    return null;
                }
                
                JsonObject issue = issueElement.getAsJsonObject();
                if (issue.has("id")) {
                    return issue.get("id").getAsString();
                } else {
                    System.out.println("No id field in issue");
                    return null;
                }
            } else {
                System.out.println("No issue field in updateIssue");
                return null;
            }
        } else {
            System.out.println("No updateIssue field in data");
            return null;
        }
    }
    
    private void updateProjectFields(String projectId, String issueId, String status, String issueType, String fixVersion, String startDate, String affectsVersion, String project) throws IOException {
        // This is a simplified implementation for updating project fields
        // In a real implementation, you would need to get field IDs and update them
        
        // For now, just ensure the issue is in the project
        String mutation = """
            mutation($input: AddProjectV2ItemByIdInput!) {
              addProjectV2ItemById(input: $input) {
                item {
                  id
                }
              }
            }
            """;
        
        JsonObject input = new JsonObject();
        input.addProperty("projectId", projectId);
        input.addProperty("contentId", issueId);
        
        JsonObject variables = new JsonObject();
        variables.add("input", input);
        
        executeGraphQL(mutation, variables);
    }
    
    private String getCellValue(Row row, Map<String, Integer> columnMap, String columnName) {
        Integer colIndex = columnMap.get(columnName.toLowerCase());
        if (colIndex == null) return "";
        
        Cell cell = row.getCell(colIndex);
        return getCellValueAsString(cell);
    }
    
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
                    return dateFormat.format(date);
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
    
    private String mapIssueTypeLabel(String issueType) {
        if (issueType == null || issueType.isEmpty()) return "";
        
        switch (issueType.toLowerCase()) {
            case "bug":
            case "bug report":
                return "bug";
            case "feature":
            case "feature request":
            case "improvement":
                return "enhancement";
            case "task":
                return "task";
            case "documentation":
                return "documentation";
            default:
                return "enhancement";
        }
    }
    
    private String populateImprovementFormTemplate(String title, String project, String issueType, String priority, 
            String affectsVersion, String fixVersion, String components, String moduleInfo, 
            String assignee, String description, String linkedIssues, String attachments, String resolution) {
        
        StringBuilder body = new StringBuilder();
        
        // Application field (from improvement_form.yml)
        body.append("**Application:** ");
        body.append(project != null && !project.isEmpty() ? project : "No response");
        body.append("\n\n");

        // Priority field (from improvement_form.yml)
        body.append("**Priority:** ");
        body.append(priority != null && !priority.isEmpty() ? priority : "No response");
        body.append("\n\n");
        
        // Affects Version/s field (from improvement_form.yml)
        body.append("**Affects Version/s:** ");
        body.append(affectsVersion != null && !affectsVersion.isEmpty() ? affectsVersion : "No response");
        body.append("\n\n");
        
        // Fix Version/s field (from improvement_form.yml)
        body.append("**Fix Version/s:** ");
        body.append(fixVersion != null && !fixVersion.isEmpty() ? fixVersion : "No response");
        body.append("\n\n");
        
        // Component/s field (from improvement_form.yml)
        body.append("**Component/s:** ");
        body.append(components != null && !components.isEmpty() ? components : "No response");
        body.append("\n\n");
        
        // Module Info field (from improvement_form.yml)
        body.append("**Module Info:** ");
        body.append(moduleInfo != null && !moduleInfo.isEmpty() ? moduleInfo : "No response");
        body.append("\n\n");
        
        // Reporter field (from improvement_form.yml)
        body.append("**Reporter:** ");
        body.append(assignee != null && !assignee.isEmpty() ? assignee : "No response");
        body.append("\n\n");
        
        // Description field (from improvement_form.yml)
        body.append("**Description:** ");
        body.append(description != null && !description.isEmpty() ? description : "No response");
        body.append("\n\n");
        
        // Linked Issues field (from improvement_form.yml)
        body.append("**Linked Issues:** ");
        body.append(linkedIssues != null && !linkedIssues.isEmpty() ? linkedIssues : "No response");
        body.append("\n\n");
        
        // Attachments field (from improvement_form.yml)
        body.append("**Attachments:** ");
        body.append(attachments != null && !attachments.isEmpty() ? attachments : "No response");
        body.append("\n\n");
        
        // Resolution field
        body.append("**Resolution:** ");
        body.append(resolution != null && !resolution.isEmpty() ? resolution : "No response");
        body.append("\n\n");
        
        return body.toString();
    }
    
    private JsonObject getRepoInfo(String owner, String repo, String projectName) throws IOException {
        String query = String.format("""
        	    query($owner: String!, $repo: String!) {
        	      user(login: $owner) {
        	        id
        	      }
        	      repository(owner: $owner, name: $repo) {
        	        id
        	        projectsV2(first: 10) {
        	          nodes {
        	            id
        	            title
        	          }
        	        }
        	      }
        	    }
        	    """);
        
        JsonObject variables = new JsonObject();
        variables.addProperty("owner", owner);
        variables.addProperty("repo", repo);
        
        JsonObject response = executeGraphQL(query, variables);
        
        JsonObject userObj = response.getAsJsonObject("data").getAsJsonObject("user");
        JsonObject repoObj = response.getAsJsonObject("data").getAsJsonObject("repository");
        
        String repoId = repoObj.get("id").getAsString();
        String ownerId = userObj.get("id").getAsString();
        
        // Get project ID
        String projectId = null;
        for (JsonElement projectEl : repoObj.getAsJsonObject("projectsV2").getAsJsonArray("nodes")) {
            JsonObject project = projectEl.getAsJsonObject();
            if (project.get("title").getAsString().equalsIgnoreCase(projectName)) {
                projectId = project.get("id").getAsString();
                break;
            }
        }
        if (projectId == null) {
            throw new IllegalStateException("Project '" + projectName + "' not found.");
        }
        
        JsonObject result = new JsonObject();
        result.addProperty("repoId", repoId);
        result.addProperty("projectId", projectId);
        result.addProperty("ownerId", ownerId);
        return result;
    }
    
    private String[] createIssue(String repoId, String title, String body, List<String> labels, List<String> assignees, String milestone) throws IOException {
        String mutation = String.format("""
            mutation($input: CreateIssueInput!) {
              createIssue(input: $input) {
                issue {
                  id
                  number
                }
              }
            }
            """);
        
        JsonObject input = new JsonObject();
        input.addProperty("repositoryId", repoId);
        input.addProperty("title", title);
        input.addProperty("body", body);
        
        // Remove labels from initial creation to avoid GraphQL schema issues
        // Labels will be added separately after issue creation
        
        JsonObject variables = new JsonObject();
        variables.add("input", input);
        
        System.out.println("Creating issue with title: " + title);
        System.out.println("GraphQL Variables: " + variables.toString());
        
        JsonObject response = executeGraphQL(mutation, variables);
        JsonElement data = response.get("data");
        
        if (data != null && data.isJsonObject() && data.getAsJsonObject().has("createIssue")) {
            JsonObject issue = data.getAsJsonObject().get("createIssue").getAsJsonObject().get("issue").getAsJsonObject();
            String issueId = issue.get("id").getAsString();
            String issueNumber = issue.get("number").getAsString();
            
            System.out.println("Created issue #" + issueNumber + " with ID: " + issueId);
            
            // Add labels to the created issue
            if (!labels.isEmpty()) {
                addLabelsToIssue(issueNumber, labels);
            }
            
            return new String[]{issueId, issueNumber};
        }
        
        return null;
    }
    
    private void addCommentToIssue(String issueId, String commentBody) throws IOException {
        String mutation = """
            mutation($input: AddCommentInput!) {
              addComment(input: $input) {
                comment {
                  id
                  body
                }
              }
            }
            """;
        
        JsonObject input = new JsonObject();
        input.addProperty("subjectId", issueId);
        input.addProperty("body", commentBody);
        
        JsonObject variables = new JsonObject();
        variables.add("input", input);
        
        System.out.println("Adding comment to issue with ID: " + issueId);
        
        JsonObject response = executeGraphQL(mutation, variables);
        
        System.out.println("Comment creation response: " + response);
        
        if (response != null) {
            if (response.has("errors")) {
                JsonArray errors = response.getAsJsonArray("errors");
                System.out.println("GraphQL Errors in comment creation: " + errors.toString());
            }
            
            if (response.has("data")) {
                JsonElement data = response.get("data");
                if (data.isJsonObject() && data.getAsJsonObject().has("addComment")) {
                    JsonObject comment = data.getAsJsonObject().get("addComment").getAsJsonObject().get("comment").getAsJsonObject();
                    String commentId = comment.get("id").getAsString();
                    System.out.println("Successfully added comment with ID: " + commentId + " to issue with ID: " + issueId);
                } else {
                    System.out.println("Failed to add comment - no addComment in response. Data: " + data.toString());
                }
            } else {
                System.out.println("Failed to add comment - no data in response. Response: " + response.toString());
            }
        } else {
            System.out.println("Failed to add comment - null response");
        }
    }
    
    private void addSecondIssueBody(String issueId, String issueNumber, String comments) throws IOException {
        // Create a comment that looks like a body update
        String commentMutation = """
            mutation($input: AddCommentInput!) {
              addComment(input: $input) {
                clientMutationId
              }
            }
            """;
        
        JsonObject commentInput = new JsonObject();
        commentInput.addProperty("subjectId", issueId);
        
        // Create comment content that looks like a body update
        StringBuilder commentBody = new StringBuilder();
        //commentBody.append("**Additional Comments**\n");
        commentBody.append(comments);
        commentBody.append("\n\n");
        
        commentInput.addProperty("body", commentBody.toString());
        
        JsonObject commentVariables = new JsonObject();
        commentVariables.add("input", commentInput);
        
        System.out.println("Adding comment as second body to issue #" + issueNumber);
        System.out.println("Comment content: " + commentBody.toString());
        
        JsonObject commentResponse = executeGraphQL(commentMutation, commentVariables);
        
        if (commentResponse != null && commentResponse.has("data")) {
            JsonElement data = commentResponse.get("data");
            if (data.isJsonObject() && data.getAsJsonObject().has("addComment")) {
                System.out.println("Successfully added second body comment to issue #" + issueNumber);
            } else {
                System.out.println("Failed to add second body comment - no addComment in response. Response: " + commentResponse.toString());
            }
        } else {
            System.out.println("Failed to add second body comment - no data in response. Response: " + commentResponse.toString());
        }
    }
    
    private void closeIssue(String issueNumber) throws IOException {
        String mutation = """
            mutation($input: CloseIssueInput!) {
              closeIssue(input: $input) {
                issue {
                  id
                  number
                  state
                }
              }
            }
            """;
        
        JsonObject input = new JsonObject();
        input.addProperty("issueId", issueNumber);
        
        JsonObject variables = new JsonObject();
        variables.add("input", input);
        
        System.out.println("Closing issue #" + issueNumber);
        
        JsonObject response = executeGraphQL(mutation, variables);
        
        if (response != null && response.has("data")) {
            JsonElement data = response.get("data");
            if (data.isJsonObject() && data.getAsJsonObject().has("closeIssue")) {
                System.out.println("Successfully closed issue #" + issueNumber);
            }
        }
    }
    
    private void addLabelsToIssue(String issueNumber, List<String> labels) throws IOException {
        if (labels == null || labels.isEmpty()) return;
        
        // Use REST API to add labels by name
        String url = String.format("https://api.github.com/repos/bk71-cdc/github-issues/issues/%s/labels", issueNumber);
        
        JsonArray labelsArray = new JsonArray();
        for (String label : labels) {
            labelsArray.add(label);
        }
        
        RequestBody body = RequestBody.create(labelsArray.toString(), MediaType.get("application/json"));
        
        Request request = new Request.Builder()
                .url(url)
                .header("Authorization", "Bearer " + GITHUB_TOKEN)
                .header("Accept", "application/vnd.github.v3+json")
                .post(body)
                .build();
        
        System.out.println("Adding labels to issue #" + issueNumber + ": " + labels);
        
        try (Response response = client.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                String errorBody = response.body() != null ? response.body().string() : "No error body";
                System.err.println("Failed to add labels: " + response.code() + " - " + errorBody);
            } else {
                System.out.println("Successfully added labels to issue #" + issueNumber);
            }
        }
    }
    
        
    private void addIssueToProject(String projectId, String issueId, String status, String issueType, String fixVersion, String startDate, String affectsVersion, String project) throws IOException {
        // This is a simplified implementation
        // In a real implementation, you would need to get the field IDs and update them
        
        String mutation = String.format("""
            mutation($input: AddProjectV2ItemByIdInput!) {
              addProjectV2ItemById(input: $input) {
                item {
                  id
                }
              }
            }
            """);
        
        JsonObject input = new JsonObject();
        input.addProperty("projectId", projectId);
        input.addProperty("contentId", issueId);
        
        JsonObject variables = new JsonObject();
        variables.add("input", input);
        
        executeGraphQL(mutation, variables);
    }
    
    private JsonObject executeGraphQL(String query) throws IOException {
        return executeGraphQL(query, null);
    }
    
    private JsonObject executeGraphQL(String query, JsonObject variables) throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty("query", query);
        if (variables != null) {
            payload.add("variables", variables);
        }

        RequestBody body = RequestBody.create(payload.toString(), MediaType.get("application/json"));

        Request request = new Request.Builder()
                .url(GITHUB_API_URL)
                .header("Authorization", "Bearer " + GITHUB_TOKEN)
                .post(body)
                .build();

        try (Response response = client.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                String errorBody = response.body() != null ? response.body().string() : "No error body";
                throw new IOException("Unexpected code: " + response.code() + " - " + errorBody);
            }

            String responseBody = response.body() != null ? response.body().string() : null;
            System.out.println("GraphQL Response Body: " + responseBody);
            
            if (responseBody == null || responseBody.trim().isEmpty()) {
                throw new IOException("Empty response body from GitHub API");
            }
            
            return JsonParser.parseString(responseBody).getAsJsonObject();
        }
    }
}