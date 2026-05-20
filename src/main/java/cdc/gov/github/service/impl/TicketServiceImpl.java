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
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.ThreadLocalRandom;

@Service
public class TicketServiceImpl implements TicketService {
    
    private static final String GITHUB_API_URL = "https://api.github.com/graphql";
    private static final String GITHUB_TOKEN = "";

    private static final String OWNER = "cdcent";
    //private static final String REPO = "NCHHSTP-DHP-HSB-EHARS-SANDBOX-GERRY";
    private static final String REPO = "NCHHSTP-DHP-HSB-ADT-SANDBOX-BINTA";

/*    private static final String OWNER = "bk71-cdc";
    private static final String REPO = "github-issues";*/


    private final OkHttpClient client = new OkHttpClient.Builder()
            .connectTimeout(30, java.util.concurrent.TimeUnit.SECONDS)
            .readTimeout(60, java.util.concurrent.TimeUnit.SECONDS)
            .writeTimeout(60, java.util.concurrent.TimeUnit.SECONDS)
            .connectionPool(new okhttp3.ConnectionPool(10, 5, java.util.concurrent.TimeUnit.MINUTES))
            .build();
    private final Gson gson = new GsonBuilder().setPrettyPrinting().create();
    private String repositoryId = "";
    private String ownerId = "";
    
    // Rate limit tracking
    private int rateLimitRemaining = -1;
    private long rateLimitResetTime = -1;
    
    // Request tracking for debugging
    private final AtomicInteger requestCounter = new AtomicInteger(0);
    private static final int CONNECTION_RESET_INTERVAL = 50; // Reset connection pool every 50 requests
    private static final int BATCH_SIZE = 50; // Process in batches of 50
    private static final int BATCH_DELAY_MS = 2000; // 2 second delay between batches

    /**
     * Comprehensive GraphQL response validation utility
     */
    private static class GraphQLResponseValidator {
        private final JsonObject response;
        private final String requestId;
        private final Gson gson;

        public GraphQLResponseValidator(JsonObject response, String requestId, Gson gson) {
            this.response = response;
            this.requestId = requestId;
            this.gson = gson;
        }

        /**
         * Validate complete GraphQL response structure
         * @throws IOException with detailed error message if validation fails
         */
        public void validate() throws IOException {
            validateResponseNotNull();
            validateErrors();
            validateDataNotNull();
        }

        /**
         * Validate createIssue mutation response specifically
         * @throws IOException with detailed error message if validation fails
         */
        public void validateCreateIssueResponse() throws IOException {
            validate();
            
            JsonElement data = response.get("data");
            JsonObject dataObj = data.getAsJsonObject();
            
            if (!dataObj.has("createIssue")) {
                throw new IOException(buildErrorMessage("Response missing 'createIssue' field"));
            }
            
            JsonElement createIssueElement = dataObj.get("createIssue");
            
            if (createIssueElement == null || createIssueElement.isJsonNull()) {
                throw new IOException(buildErrorMessage("createIssue is null - mutation may have failed silently"));
            }
            
            if (!createIssueElement.isJsonObject()) {
                throw new IOException(buildErrorMessage("createIssue is not an object"));
            }
            
            JsonObject createIssueObj = createIssueElement.getAsJsonObject();
            
            // Check if createIssue is empty object (partial failure)
            if (createIssueObj.size() == 0) {
                throw new IOException(buildErrorMessage("createIssue is empty object {} - indicates partial failure or server-side rejection"));
            }
            
            if (!createIssueObj.has("issue")) {
                throw new IOException(buildErrorMessage("createIssue missing 'issue' field - mutation succeeded but didn't return issue data"));
            }
            
            JsonElement issueElement = createIssueObj.get("issue");
            
            if (issueElement == null || issueElement.isJsonNull()) {
                throw new IOException(buildErrorMessage("issue is null - issue creation may have failed or been rejected"));
            }
            
            if (!issueElement.isJsonObject()) {
                throw new IOException(buildErrorMessage("issue is not an object"));
            }
            
            JsonObject issue = issueElement.getAsJsonObject();
            
            if (issue.size() == 0) {
                throw new IOException(buildErrorMessage("issue object is empty - indicates data corruption or server error"));
            }
            
            if (!issue.has("id") || !issue.has("number")) {
                throw new IOException(buildErrorMessage("issue missing required fields (id or number). Issue data: " + gson.toJson(issue)));
            }
        }

        private void validateResponseNotNull() throws IOException {
            if (response == null) {
                throw new IOException(buildErrorMessage("Response is null"));
            }
        }

        private void validateErrors() throws IOException {
            if (response.has("errors")) {
                JsonArray errors = response.getAsJsonArray("errors");
                StringBuilder errorMessages = new StringBuilder();
                for (JsonElement error : errors) {
                    if (error != null && error.isJsonObject()) {
                        JsonObject errorObj = error.getAsJsonObject();
                        if (errorObj.has("message")) {
                            errorMessages.append(errorObj.get("message").getAsString()).append("; ");
                        }
                        if (errorObj.has("type")) {
                            errorMessages.append("[Type: ").append(errorObj.get("type").getAsString()).append("] ");
                        }
                        if (errorObj.has("path")) {
                            errorMessages.append("[Path: ").append(errorObj.get("path")).append("] ");
                        }
                    }
                }
                throw new IOException(buildErrorMessage("GraphQL errors: " + errorMessages.toString()));
            }
        }

        private void validateDataNotNull() throws IOException {
            JsonElement data = response.get("data");
            if (data == null || data.isJsonNull()) {
                throw new IOException(buildErrorMessage("Response data is null - indicates complete failure"));
            }
            if (!data.isJsonObject()) {
                throw new IOException(buildErrorMessage("Response data is not an object"));
            }
        }

        private String buildErrorMessage(String message) {
            return String.format("[Request ID: %s] %s. Full response: %s", 
                requestId, message, gson.toJson(response));
        }
    }

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
                    // Add throttling delay between requests to avoid rate limiting
                    if (i > 0) {
                        Thread.sleep(200); // 200ms delay between requests
                    }
                    
                    String issueTitle = "Empty Issue #" + (i + 1);
                    String issueBody = populateImprovementFormTemplate(
                        issueTitle, "eHARS", "Task", "Average", 
                        "Release 4.20", "Release 4.21", "All Document Import", "No Module", 
                        "", "This is an auto-generated empty issue.", "", "", "", "", "", "", "", "", "");
                    
                    // Call GitHub API to create issue
                    String[] issueResult = createIssue(repositoryId, issueTitle, issueBody, 
                        Arrays.asList("improvement"), new ArrayList<>(), "");
                    
                    if (issueResult != null && issueResult.length == 2) {
                        String issueId = issueResult[0];
                        String issueNumber = issueResult[1];
                        createdIssueIds.add(issueId);
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




    public LoadTicketsResponse processAdtFromExcel(String fileName) {
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

            // Process the Excel file for ADT entries only
            List<String> createdIssueIds = processAdtExcelFile(excelStream, response);

            response.setTotalCount(createdIssueIds.size());
            response.setSuccessCount(createdIssueIds.size());
            response.setFailureCount(0);
            response.setIssueIds(createdIssueIds);
            response.setMessage("Successfully created " + createdIssueIds.size() + " ADT issues from " + fileName);

        } catch (Exception e) {
            response.setMessage("Error processing ADT tickets from Excel: " + e.getMessage());
            response.addError("Exception: " + e.getMessage());
            response.setTotalCount(0);
            response.setSuccessCount(0);
            response.setFailureCount(1);
        }

        return response;
    }

    private void initializeRepositoryInfo() throws IOException {
        if (repositoryId.isEmpty()) {
            try {
                JsonObject repoInfo = getRepoInfo(OWNER, REPO);
                repositoryId = repoInfo.get("repoId").getAsString();
                ownerId = repoInfo.get("ownerId").getAsString();
            } catch (Exception e) {
                throw e;
            }
        }
    }
    
    private InputStream getExcelInputStream(String fileName) {
        try {
            // First try to load from resources
            ClassPathResource resource = new ClassPathResource(fileName);
            if (resource.exists()) {
                return resource.getInputStream();
            }
            
            // Try as absolute path
            File file = new File(fileName);
            if (file.exists()) {
                return new FileInputStream(file);
            }
            
            // Try as relative path from project root
            File relativeFile = new File(System.getProperty("user.dir"), fileName);
            if (relativeFile.exists()) {
                return new FileInputStream(relativeFile);
            }
            
            return null;
        } catch (Exception e) {
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
                } else {
                    // Try XLSX first, if it fails, fall back to XLS
                    try {
                        workbook = new XSSFWorkbook(excelStream);
                    } catch (Exception xlsxException) {
                        // Reset stream and try XLS
                        excelStream.reset();
                        workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
                    }
                }
            } else {
                // If stream doesn't support mark, try XLSX first, then XLS
                try {
                    workbook = new XSSFWorkbook(excelStream);
                } catch (Exception xlsxException) {
                    // Create a fresh stream for XLS
                    workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
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
            
            // Process data rows, starting from row 5 (0-indexed is 4)
            for (int rowNum = 4; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;
                
                // Extract key field for title prefix
                String excelIssueId = getCellValue(row, columnMap, "id");
                String excelTicketId = getCellValue(row, columnMap, "ticket id");
                String excelKey = getCellValue(row, columnMap, "key");
                
                // Use the first non-empty key field found
                String keyPrefix = "";
                if (excelIssueId != null && !excelIssueId.isEmpty()) {
                    keyPrefix = excelIssueId;
                } else if (excelTicketId != null && !excelTicketId.isEmpty()) {
                    keyPrefix = excelTicketId;
                } else if (excelKey != null && !excelKey.isEmpty()) {
                    keyPrefix = excelKey;
                }
                
                // Extract data from key columns
                String title = getCellValue(row, columnMap, "summary");
                if (title.isEmpty()) {
                    continue;
                }
                
                // Prepend key to title if key exists
                if (!keyPrefix.isEmpty()) {
                    title = keyPrefix + ": " + title;
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
                String created = getCellValue(row, columnMap, "created");
                String updated = getCellValue(row, columnMap, "updated");
                String customerName = getCellValue(row, columnMap, "customer name");
                String helpDeskNumbers = getCellValue(row, columnMap, "help desk number(s)");
                String duplicateCqId = getCellValue(row, columnMap, "duplicate cq id");
                String resolved = getCellValue(row, columnMap, "resolved");
                
                // Check if Module Info is "ADT" - if so, create empty issue
                if ("ADT".equalsIgnoreCase(moduleInfo)) {
                    try {
                        String emptyIssueTitle = !keyPrefix.isEmpty() ? keyPrefix + ": Empty Issue" : "Empty Issue";
                        String emptyIssueBody = ""; // Only title, no other template data

                        String[] issueResult = createIssue(repositoryId, emptyIssueTitle, emptyIssueBody,
                            Arrays.asList("improvement"), new ArrayList<>(), "");

                        if (issueResult != null && issueResult.length == 2) {
                            String issueId = issueResult[0];
                            String issueNumber = issueResult[1];
                            createdIssueIds.add("Issue #" + issueNumber);

                            // Close issue if status is "closed" from Excel
                            if ("closed".equalsIgnoreCase(status)) {
                                try {
                                    closeIssue(issueId);
                                } catch (Exception e) {
                                    // Silently ignore close failures
                                }
                            }
                        } else {
                            response.addError("Failed to create empty issue for ADT module in row " + rowNum);
                        }
                    } catch (Exception e) {
                        response.addError("Error creating empty issue for ADT module in row " + rowNum + ": " + e.getMessage());
                    }
                    continue; // Skip normal processing for this row
                }

                // Map issue type to label
                String mappedIssueType = mapIssueTypeLabel(issueType);
                
                // Set default project since not in Excel headers
                String projectValue = project != null && !project.isEmpty() ? project : "eHARS";
                
                // Create issue body using improvement_form.yml structure
                String body = populateImprovementFormTemplate(
                    title, projectValue, mappedIssueType, priority, 
                    affectsVersion, fixVersion, components, moduleInfo, 
                    assignee, description, linkedIssues, attachment, resolution,
                    created, updated, customerName, helpDeskNumbers, duplicateCqId, resolved);
                
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
                
                String[] issueResult;
                try {
                    // Add throttling delay between requests to avoid rate limiting
                    if (rowNum > 4) { // Skip delay for first request
                        Thread.sleep(200); // 200ms delay between requests
                    }
                    
                    issueResult = createIssue(repositoryId, title, body, labels, assignees, "");
                } catch (Exception e) {
                    response.addError("Failed to create issue for row " + rowNum + ": " + title + " - Error: " + e.getMessage());
                    continue;
                }
                
                if (issueResult != null && issueResult.length == 2) {
                    String issueId = issueResult[0];
                    String issueNumber = issueResult[1];
                    
                    // Close issue if status is "closed"
                    if ("closed".equalsIgnoreCase(status)) {
                        try {
                            closeIssue(issueId);
                        } catch (Exception e) {
                            // Silently ignore close failures
                        }
                    }
                    
                    createdIssueIds.add("Issue #" + issueNumber);
                    
                    // Add second issue body with comments as the very last step
                    if (comments != null && !comments.isEmpty()) {
                        try {
                            addSecondIssueBody(issueId, issueNumber, comments);
                        } catch (Exception e) {
                            // Silently ignore comment addition failures
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
    
    private List<String> processAdtExcelFile(InputStream excelStream, LoadTicketsResponse response) throws IOException {
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
                } else {
                    // Try XLSX first, if it fails, fall back to XLS
                    try {
                        workbook = new XSSFWorkbook(excelStream);
                    } catch (Exception xlsxException) {
                        // Reset stream and try XLS
                        excelStream.reset();
                        workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
                    }
                }
            } else {
                // If stream doesn't support mark, try XLSX first, then XLS
                try {
                    workbook = new XSSFWorkbook(excelStream);
                } catch (Exception xlsxException) {
                    // Create a fresh stream for XLS
                    workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
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

            // Process data rows, starting from row 5 (0-indexed is 4)
            for (int rowNum = 4; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;

                // Extract key field for title prefix
                String excelIssueId = getCellValue(row, columnMap, "id");
                String excelTicketId = getCellValue(row, columnMap, "ticket id");
                String excelKey = getCellValue(row, columnMap, "key");

                // Use the first non-empty key field found
                String keyPrefix = "";
                if (excelIssueId != null && !excelIssueId.isEmpty()) {
                    keyPrefix = excelIssueId;
                } else if (excelTicketId != null && !excelTicketId.isEmpty()) {
                    keyPrefix = excelTicketId;
                } else if (excelKey != null && !excelKey.isEmpty()) {
                    keyPrefix = excelKey;
                }

                // Extract data from key columns
                String title = getCellValue(row, columnMap, "summary");
                if (title.isEmpty()) {
                    continue;
                }

                // Prepend key to title if key exists
                if (!keyPrefix.isEmpty()) {
                    title = keyPrefix + ": " + title;
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
                String created = getCellValue(row, columnMap, "created");
                String updated = getCellValue(row, columnMap, "updated");
                String customerName = getCellValue(row, columnMap, "customer name");
                String helpDeskNumbers = getCellValue(row, columnMap, "help desk number(s)");
                String duplicateCqId = getCellValue(row, columnMap, "duplicate cq id");
                String resolved = getCellValue(row, columnMap, "resolved");

                // Check if Module Info is "ADT" - if so, create full issue, otherwise skip
                if (!"ADT".equalsIgnoreCase(moduleInfo)) {
                    continue;
                }

                // Map issue type to label
                String mappedIssueType = mapIssueTypeLabel(issueType);

                // Set default project since not in Excel headers
                String projectValue = project != null && !project.isEmpty() ? project : "eHARS";

                // Create issue body using improvement_form.yml structure
                String body = populateImprovementFormTemplate(
                    title, projectValue, mappedIssueType, priority,
                    affectsVersion, fixVersion, components, moduleInfo,
                    assignee, description, linkedIssues, attachment, resolution,
                    created, updated, customerName, helpDeskNumbers, duplicateCqId, resolved);

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

                // Create full issue for ADT entries
                try {
                    // Add throttling delay between requests to avoid rate limiting
                    if (rowNum > 4) {
                        Thread.sleep(200); // 200ms delay between requests
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
                                // Silently ignore close failures
                            }
                        }

                        createdIssueIds.add("Issue #" + issueNumber);

                        // Add second issue body with comments as very last step
                        if (comments != null && !comments.isEmpty()) {
                            try {
                                addSecondIssueBody(issueId, issueNumber, comments);
                            } catch (Exception e) {
                                // Silently ignore comment addition failures
                            }
                        }
                    } else {
                        response.addError("Failed to create issue for ADT row " + rowNum + ": " + title);
                    }
                } catch (Exception e) {
                    response.addError("Error creating issue for ADT row " + rowNum + ": " + e.getMessage());
                }
            }

            workbook.close();
        } catch (Exception e) {
            throw new IOException("Failed to process ADT Excel file: " + e.getMessage(), e);
        }

        return createdIssueIds;
    }

    private List<String> processExcelFileForUpdates(InputStream excelStream, LoadTicketsResponse response) throws IOException {
        List<String> updatedIssueIds = new ArrayList<>();
        
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
                } else {
                    // Try XLSX first, if it fails, fall back to XLS
                    try {
                        workbook = new XSSFWorkbook(excelStream);
                    } catch (Exception xlsxException) {
                        // Reset stream and try XLS
                        excelStream.reset();
                        workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
                    }
                }
            } else {
                // If stream doesn't support mark, try XLSX first, then XLS
                try {
                    workbook = new XSSFWorkbook(excelStream);
                } catch (Exception xlsxException) {
                    // Create a fresh stream for XLS
                    workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
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
            
            // Process data rows, starting from row 5 (0-indexed is 4)
            for (int rowNum = 4; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;
                
                // Extract issue ID (key field like EHARS-600)
                String issueId = getCellValue(row, columnMap, "id");
                String ticketId = getCellValue(row, columnMap, "ticket id");
                String key = getCellValue(row, columnMap, "key");
                
                // Try to find issue number from any of these columns
                String issueNumber = extractIssueNumber(issueId);
                String originalKey = issueId; // Store the original key value
                if (issueNumber.isEmpty()) {
                    issueNumber = extractIssueNumber(ticketId);
                    originalKey = ticketId; // Use ticketId as original key if issueId was empty
                }
                if (issueNumber.isEmpty()) {
                    issueNumber = extractIssueNumber(key);
                    originalKey = key; // Use key as original key if others were empty
                }
                
                if (issueNumber.isEmpty()) {
                    response.addError("No valid issue ID found in row " + rowNum + ". Expected format like EHARS-600");
                    continue;
                }
                
                // Extract other data from key columns
                String title = getCellValue(row, columnMap, "summary");
                
                // Prepend key to title if key exists
                if (!title.isEmpty() && originalKey != null && !originalKey.isEmpty()) {
                    title = originalKey + ": " + title;
                }
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
                String created = getCellValue(row, columnMap, "created");
                String updated = getCellValue(row, columnMap, "updated");
                String customerName = getCellValue(row, columnMap, "customer name");
                String helpDeskNumbers = getCellValue(row, columnMap, "help desk number(s)");
                String duplicateCqId = getCellValue(row, columnMap, "duplicate cq id");
                String resolved = getCellValue(row, columnMap, "resolved");
                
                // Map issue type to label
                String mappedIssueType = mapIssueTypeLabel(issueType);
                
                // Set default project since not in Excel headers
                String projectValue = project != null && !project.isEmpty() ? project : "eHARS";
                
                // Create issue body using improvement_form.yml structure
                String body = populateImprovementFormTemplate(
                    title, projectValue, mappedIssueType, priority, 
                    affectsVersion, fixVersion, components, moduleInfo, 
                    assignee, description, linkedIssues, attachment, resolution,
                    created, updated, customerName, helpDeskNumbers, duplicateCqId, resolved);
                
                // Update existing GitHub issue
                try {
                    String updatedIssueId = updateIssue(repositoryId, issueNumber, title, body, mappedIssueType, assignee);
                    if (updatedIssueId != null && !updatedIssueId.isEmpty()) {
                        updatedIssueIds.add("Issue #" + issueNumber);
                        
                        // Handle issue status - close if "closed", otherwise keep open
                        String status = getCellValue(row, columnMap, "status");
                        if ("closed".equalsIgnoreCase(status)) {
                            try {
                                closeIssue(updatedIssueId);
                            } catch (Exception e) {
                                // Silently ignore close failures
                            }
                        }
                        
                        // Add second issue body with comments as the very last step
                        if (comments != null && !comments.isEmpty()) {
                            try {
                                addSecondIssueBody(updatedIssueId, issueNumber, comments);
                            } catch (Exception e) {
                                // Silently ignore comment addition failures
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
        } catch (Exception e) {
            response.addError("Error processing Excel file: " + e.getMessage());
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
    
    private String getIssueNodeId(String repoId, String issueNumber) throws IOException {
        String query = """
            query($owner: String!, $repo: String!, $number: Int!) {
              repository(owner: $owner, name: $repo) {
                issue(number: $number) {
                  id
                }
              }
            }
            """;
        
        JsonObject variables = new JsonObject();
        variables.addProperty("owner", OWNER);
        variables.addProperty("repo", REPO);
        variables.addProperty("number", Integer.parseInt(issueNumber));
        
        JsonObject response = executeGraphQL(query, variables);
        
        if (response == null) {
            return null;
        }
        
        if (response.has("errors")) {
            JsonArray errors = response.getAsJsonArray("errors");
            System.err.println("GraphQL errors getting node ID for issue #" + issueNumber + ": " + errors);
            return null;
        }
        
        JsonElement data = response.get("data");
        if (data == null || data.isJsonNull()) {
            return null;
        }
        
        JsonObject repository = data.getAsJsonObject().getAsJsonObject("repository");
        if (repository == null || repository.isJsonNull()) {
            return null;
        }
        
        JsonObject issue = repository.getAsJsonObject("issue");
        if (issue == null || issue.isJsonNull()) {
            return null;
        }
        
        return issue.get("id").getAsString();
    }
    
    private String updateIssue(String repoId, String issueNumber, String title, String body, String issueType, String assignee) throws IOException {
        // First, get the issue's node ID
        String issueNodeId = getIssueNodeId(repoId, issueNumber);
        if (issueNodeId == null) {
            System.err.println("Could not get node ID for issue #" + issueNumber);
            return null;
        }
        
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
        input.addProperty("id", issueNodeId);
        
        if (title != null && !title.isEmpty()) {
            input.addProperty("title", title);
        }
        
        if (body != null && !body.isEmpty()) {
            input.addProperty("body", body);
        }
        
        JsonObject variables = new JsonObject();
        variables.add("input", input);
        
        JsonObject response = executeGraphQL(mutation, variables);
        
        if (response == null) {
            System.err.println("GraphQL response is null for issue #" + issueNumber);
            return null;
        }
        
        // Check for errors in GraphQL response
        if (response.has("errors")) {
            JsonArray errors = response.getAsJsonArray("errors");
            System.err.println("GraphQL errors for issue #" + issueNumber + ": " + errors);
            return null;
        }
        
        JsonElement data = response.get("data");
        if (data == null || data.isJsonNull()) {
            return null;
        }
        
        if (data.isJsonObject() && data.getAsJsonObject().has("updateIssue")) {
            JsonElement updateIssueElement = new Gson().fromJson(data.getAsJsonObject().get("updateIssue").toString(), JsonElement.class);
            if (updateIssueElement.isJsonNull()) {
                return null;
            }
            
            JsonObject updateIssue = updateIssueElement.getAsJsonObject();
            if (updateIssue.has("issue")) {
                JsonElement issueElement = updateIssue.get("issue");
                if (issueElement.isJsonNull()) {
                    return null;
                }
                
                JsonObject issue = issueElement.getAsJsonObject();
                if (issue.has("id")) {
                    return issue.get("id").getAsString();
                } else {
                    return null;
                }
            } else {
                return null;
            }
        } else {
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
            case "new feature":
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
            String assignee, String description, String linkedIssues, String attachments, String resolution,
            String created, String updated, String customerName, String helpDeskNumbers, String duplicateCqId, String resolved) {
        
        StringBuilder body = new StringBuilder();
        
        // Application field (from improvement_form.yml)
        body.append("**Application:** ");
        body.append(project != null && !project.isEmpty() ? project : "No response");
        body.append("\n\n");

        // Priority field (from improvement_form.yml)
        body.append("**Priority:** ");
        body.append(priority != null && !priority.isEmpty() ? priority : "No response");
        body.append("\n\n");

        // Description field (from improvement_form.yml)
        body.append("**Description:** ");
        body.append(description != null && !description.isEmpty() ? description : "No response");
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
        
        // Created field
        body.append("**Created:** ");
        body.append(created != null && !created.isEmpty() ? created : "No response");
        body.append("\n\n");
        
        // Updated field
        body.append("**Updated:** ");
        body.append(updated != null && !updated.isEmpty() ? updated : "No response");
        body.append("\n\n");
        
        // Customer Name field
        body.append("**Customer Name:** ");
        body.append(customerName != null && !customerName.isEmpty() ? customerName : "No response");
        body.append("\n\n");
        
        // Help Desk Number(s) field
        body.append("**Help Desk Number(s):** ");
        body.append(helpDeskNumbers != null && !helpDeskNumbers.isEmpty() ? helpDeskNumbers : "No response");
        body.append("\n\n");
        
        // Duplicate CQ ID field
        body.append("**Duplicate ID:** ");
        if (duplicateCqId != null && !duplicateCqId.isEmpty()) {
            body.append("#").append(duplicateCqId);
        } else {
            body.append("No response");
        }
        body.append("\n\n");
        
        // Resolved field
        body.append("**Resolution Date:** ");
        body.append(resolved != null && !resolved.isEmpty() ? resolved : "No response");
        body.append("\n\n");
        
        return body.toString();
    }
    
    private JsonObject getRepoInfo(String owner, String repo) throws IOException {
        String query = String.format("""
        	    query($owner: String!, $repo: String!) {
        	      repository(owner: $owner, name: $repo) {
        	        id
        	        owner {
        	          id
        	        }
        	      }
        	    }
        	    """);
        
        JsonObject variables = new JsonObject();
        variables.addProperty("owner", owner);
        variables.addProperty("repo", repo);
        
        JsonObject response = executeGraphQL(query, variables);
        
        // Check for GraphQL errors
        if (response.has("errors")) {
            JsonArray errors = response.getAsJsonArray("errors");
            StringBuilder errorMessages = new StringBuilder();
            for (JsonElement error : errors) {
                if (error != null && error.isJsonObject()) {
                    JsonObject errorObj = error.getAsJsonObject();
                    if (errorObj.has("message")) {
                        errorMessages.append(errorObj.get("message").getAsString()).append("; ");
                    }
                }
            }
            throw new IOException("GraphQL errors: " + errorMessages.toString());
        }

        JsonElement data = response.get("data");
        if (data == null || data.isJsonNull()) {
            throw new IOException("No data returned from GraphQL query");
        }

        if (!data.isJsonObject()) {
            throw new IOException("Data is not a JSON object");
        }

        JsonObject dataObj = data.getAsJsonObject();
        JsonElement repoElement = dataObj.get("repository");

        if (repoElement == null || repoElement.isJsonNull()) {
            throw new IOException("Repository not found: " + owner + "/" + repo);
        }

        if (!repoElement.isJsonObject()) {
            throw new IOException("Repository element is not a JSON object");
        }

        JsonObject repoObj = repoElement.getAsJsonObject();
        JsonElement ownerElement = repoObj.get("owner");

        if (ownerElement == null || ownerElement.isJsonNull()) {
            throw new IOException("Owner not found for repository: " + owner + "/" + repo);
        }

        if (!ownerElement.isJsonObject()) {
            throw new IOException("Owner element is not a JSON object");
        }

        JsonObject ownerObj = ownerElement.getAsJsonObject();
        
        String repoId = repoObj.get("id").getAsString();
        String ownerId = ownerObj.get("id").getAsString();
        
        JsonObject result = new JsonObject();
        result.addProperty("repoId", repoId);
        result.addProperty("ownerId", ownerId);
        return result;
    }
    
    private String[] createIssue(String repoId, String title, String body, List<String> labels, List<String> assignees, String milestone) throws IOException {
        String requestId = "REQ-" + requestCounter.incrementAndGet();
        
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
        
        // Log request payload for debugging
        System.out.println("[" + requestId + "] Creating issue: " + title);
        System.out.println("[" + requestId + "] Request payload: " + gson.toJson(variables));
        
        JsonObject response = executeGraphQL(mutation, variables, requestId);
        
        // Log full response for debugging
        System.out.println("[" + requestId + "] GraphQL Response: " + gson.toJson(response));
        
        // Use comprehensive validator
        GraphQLResponseValidator validator = new GraphQLResponseValidator(response, requestId, gson);
        try {
            validator.validateCreateIssueResponse();
        } catch (IOException e) {
            System.err.println("[" + requestId + "] VALIDATION FAILED: " + e.getMessage());
            throw e;
        }
        
        // Extract issue data after validation
        JsonElement data = response.get("data");
        JsonObject dataObj = data.getAsJsonObject();
        JsonObject createIssueObj = dataObj.getAsJsonObject("createIssue");
        JsonObject issue = createIssueObj.getAsJsonObject("issue");
        
        String issueId = issue.get("id").getAsString();
        String issueNumber = issue.get("number").getAsString();
        
        System.out.println("[" + requestId + "] Successfully created issue #" + issueNumber);
        
        // Add labels to the created issue
        if (!labels.isEmpty()) {
            addLabelsToIssue(issueNumber, labels);
        }
        
        return new String[]{issueId, issueNumber};
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
        
        JsonObject response = executeGraphQL(mutation, variables);
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
        

        JsonObject commentResponse = executeGraphQL(commentMutation, commentVariables);
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
        
        JsonObject response = executeGraphQL(mutation, variables);
    }
    
    private void addLabelsToIssue(String issueNumber, List<String> labels) throws IOException {
        if (labels == null || labels.isEmpty()) return;
        
        // Use REST API to add labels by name
        //String url = String.format("https://api.github.com/repos/cdcent/NCHHSTP-DHP-HSB-EHARS-SANDBOX-GERRY/issues/%s/labels", issueNumber);
        String url = String.format("https://api.github.com/repos/cdcent/NCHHSTP-DHP-HSB-ADT-SANDBOX-BINTA/issues/%s/labels", issueNumber);

     //   String url = String.format("https://api.github.com/repos/bk71-cdc/github-issues/issues/%s/labels", issueNumber);
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
        
        try (Response response = client.newCall(request).execute()) {
            // Silently ignore response
        }
    }

    private JsonObject executeGraphQL(String query) throws IOException {
        return executeGraphQL(query, null, 3);
    }
    
    private JsonObject executeGraphQL(String query, JsonObject variables) throws IOException {
        return executeGraphQL(query, variables, 3, null);
    }
    
    private JsonObject executeGraphQL(String query, JsonObject variables, String requestId) throws IOException {
        return executeGraphQL(query, variables, 3, requestId);
    }
    
    private JsonObject executeGraphQL(String query, JsonObject variables, int maxRetries) throws IOException {
        return executeGraphQL(query, variables, maxRetries, null);
    }
    
    private JsonObject executeGraphQL(String query, JsonObject variables, int maxRetries, String requestId) throws IOException {
        int attempt = 0;
        IOException lastException = null;
        
        while (attempt <= maxRetries) {
            attempt++;
            
            try {
                // Check rate limit before making request
                if (rateLimitRemaining == 0) {
                    long waitTime = rateLimitResetTime - System.currentTimeMillis() / 1000;
                    if (waitTime > 0) {
                        System.out.println("Rate limit exceeded. Waiting " + waitTime + " seconds until reset.");
                        Thread.sleep(waitTime * 1000);
                        rateLimitRemaining = -1; // Reset after waiting
                    }
                } else if (rateLimitRemaining > 0 && rateLimitRemaining < 100) {
                    System.out.println("Rate limit low: " + rateLimitRemaining + " remaining. Consider throttling.");
                }
                
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
                    // Extract rate limit headers
                    String remaining = response.header("X-RateLimit-Remaining");
                    String reset = response.header("X-RateLimit-Reset");
                    
                    if (remaining != null) {
                        rateLimitRemaining = Integer.parseInt(remaining);
                        System.out.println("Rate limit remaining: " + rateLimitRemaining);
                    }
                    
                    if (reset != null) {
                        rateLimitResetTime = Long.parseLong(reset);
                    }
                    
                    // Handle rate limit (HTTP 429)
                    if (response.code() == 429) {
                        String retryAfter = response.header("Retry-After");
                        long waitTime = retryAfter != null ? Long.parseLong(retryAfter) * 1000 : 60000;
                        
                        System.err.println("Rate limit hit (429). Waiting " + (waitTime / 1000) + " seconds before retry.");
                        
                        if (attempt <= maxRetries) {
                            Thread.sleep(waitTime);
                            continue;
                        } else {
                            throw new IOException("Rate limit exceeded after " + maxRetries + " retries");
                        }
                    }
                    
                    // Handle other HTTP errors
                    if (!response.isSuccessful()) {
                        String errorBody = response.body() != null ? response.body().string() : "No error body";
                        String errorMsg = "HTTP " + response.code() + ": " + errorBody;
                        
                        // Retry on server errors (5xx) and certain client errors
                        if ((response.code() >= 500 || response.code() == 408) && attempt <= maxRetries) {
                            long backoffTime = calculateExponentialBackoff(attempt);
                            System.err.println(errorMsg + ". Retrying in " + backoffTime + "ms (attempt " + attempt + "/" + maxRetries + ")");
                            Thread.sleep(backoffTime);
                            continue;
                        }
                        
                        throw new IOException(errorMsg);
                    }

                    String responseBody = response.body() != null ? response.body().string() : null;

                    if (responseBody == null || responseBody.trim().isEmpty()) {
                        if (attempt <= maxRetries) {
                            long backoffTime = calculateExponentialBackoff(attempt);
                            System.err.println("Empty response body. Retrying in " + backoffTime + "ms (attempt " + attempt + "/" + maxRetries + ")");
                            Thread.sleep(backoffTime);
                            continue;
                        }
                        throw new IOException("Empty response body from GitHub API after " + maxRetries + " retries");
                    }
                    
                    try {
                        JsonElement jsonElement = JsonParser.parseString(responseBody);
                        if (jsonElement == null || jsonElement.isJsonNull()) {
                            if (attempt <= maxRetries) {
                                long backoffTime = calculateExponentialBackoff(attempt);
                                System.err.println("Null JSON response. Retrying in " + backoffTime + "ms (attempt " + attempt + "/" + maxRetries + ")");
                                Thread.sleep(backoffTime);
                                continue;
                            }
                            throw new IOException("Failed to parse JSON response: null after " + maxRetries + " retries");
                        }
                        return jsonElement.getAsJsonObject();
                    } catch (JsonSyntaxException e) {
                        if (attempt <= maxRetries) {
                            long backoffTime = calculateExponentialBackoff(attempt);
                            System.err.println("JSON parse error: " + e.getMessage() + ". Retrying in " + backoffTime + "ms (attempt " + attempt + "/" + maxRetries + ")");
                            Thread.sleep(backoffTime);
                            continue;
                        }
                        throw new IOException("Failed to parse JSON response: " + e.getMessage() + ". Response body: " + responseBody);
                    }
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                throw new IOException("Request interrupted", e);
            } catch (IOException e) {
                lastException = e;
                if (attempt <= maxRetries) {
                    long backoffTime = calculateExponentialBackoff(attempt);
                    System.err.println("Request failed: " + e.getMessage() + ". Retrying in " + backoffTime + "ms (attempt " + attempt + "/" + maxRetries + ")");
                    try {
                        Thread.sleep(backoffTime);
                    } catch (InterruptedException ie) {
                        Thread.currentThread().interrupt();
                        throw new IOException("Retry interrupted", ie);
                    }
                }
            }
        }
        
        throw lastException != null ? lastException : new IOException("Failed after " + maxRetries + " retries");
    }
    
    private long calculateExponentialBackoff(int attempt) {
        // Exponential backoff: 1s, 2s, 4s, 8s, etc., with jitter
        long baseDelay = 1000L; // 1 second base
        long maxDelay = 30000L; // 30 seconds max
        long delay = (long) (baseDelay * Math.pow(2, attempt - 1));
        
        // Add jitter (±25%)
        long jitter = (long) (delay * 0.25 * (Math.random() * 2 - 1));
        delay = delay + jitter;
        
        return Math.min(delay, maxDelay);
    }


    @Override
    public LoadTicketsResponse loadTicketsADT(String fileName) {

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
            List<String> createdIssueIds = processADTExcelFile(excelStream, response);

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

    private List<String> processADTExcelFile(InputStream excelStream, LoadTicketsResponse response) throws IOException {
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
                } else {
                    // Try XLSX first, if it fails, fall back to XLS
                    try {
                        workbook = new XSSFWorkbook(excelStream);
                    } catch (Exception xlsxException) {
                        // Reset stream and try XLS
                        excelStream.reset();
                        workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
                    }
                }
            } else {
                // If stream doesn't support mark, try XLSX first, then XLS
                try {
                    workbook = new XSSFWorkbook(excelStream);
                } catch (Exception xlsxException) {
                    // Create a fresh stream for XLS
                    workbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(excelStream);
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

            // Process data rows, starting from row 5 (0-indexed is 4)
            for (int rowNum = 4; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;

                // Extract key field for title prefix
                String excelIssueId = getCellValue(row, columnMap, "id");
                String excelTicketId = getCellValue(row, columnMap, "ticket id");
                String excelKey = getCellValue(row, columnMap, "key");

                // Use the first non-empty key field found
                String keyPrefix = "";
                if (excelIssueId != null && !excelIssueId.isEmpty()) {
                    keyPrefix = excelIssueId;
                } else if (excelTicketId != null && !excelTicketId.isEmpty()) {
                    keyPrefix = excelTicketId;
                } else if (excelKey != null && !excelKey.isEmpty()) {
                    keyPrefix = excelKey;
                }

                // Extract data from key columns
                String title = getCellValue(row, columnMap, "summary");
                if (title.isEmpty()) {
                    continue;
                }

                // Prepend key to title if key exists
                if (!keyPrefix.isEmpty()) {
                    title = keyPrefix + ": " + title;
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
                String created = getCellValue(row, columnMap, "created");
                String updated = getCellValue(row, columnMap, "updated");
                String customerName = getCellValue(row, columnMap, "customer name");
                String helpDeskNumbers = getCellValue(row, columnMap, "help desk number(s)");
                String duplicateCqId = getCellValue(row, columnMap, "duplicate cq id");
                String resolved = getCellValue(row, columnMap, "resolved");

                // Check if Module Info is "ADT" - if so, create empty issue
                if ("ADT".equalsIgnoreCase(moduleInfo))
                {
                    // Map issue type to label
                    String mappedIssueType = mapIssueTypeLabel(issueType);

                    // Set default project since not in Excel headers
                    String projectValue = project != null && !project.isEmpty() ? project : "eHARS";

                    // Create issue body using improvement_form.yml structure
                    String body = populateImprovementFormTemplate(
                            title, projectValue, mappedIssueType, priority,
                            affectsVersion, fixVersion, components, moduleInfo,
                            assignee, description, linkedIssues, attachment, resolution,
                            created, updated, customerName, helpDeskNumbers, duplicateCqId, resolved);

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

                    String[] issueResult;
                    try {
                        // Add throttling delay between requests to avoid rate limiting
                        if (rowNum > 4) { // Skip delay for first request
                            Thread.sleep(200); // 200ms delay between requests
                        }

                        issueResult = createIssue(repositoryId, title, body, labels, assignees, "");
                    } catch (Exception e) {
                        response.addError("Failed to create issue for row " + rowNum + ": " + title + " - Error: " + e.getMessage());
                        continue;
                    }

                    if (issueResult != null && issueResult.length == 2) {
                        String issueId = issueResult[0];
                        String issueNumber = issueResult[1];

                        // Close issue if status is "closed"
                        if ("closed".equalsIgnoreCase(status)) {
                            try {
                                closeIssue(issueId);
                            } catch (Exception e) {
                                // Silently ignore close failures
                            }
                        }

                        createdIssueIds.add("Issue #" + issueNumber);

                        // Add second issue body with comments as the very last step
                        if (comments != null && !comments.isEmpty()) {
                            try {
                                addSecondIssueBody(issueId, issueNumber, comments);
                            } catch (Exception e) {
                                // Silently ignore comment addition failures
                            }
                        }
                    } else {
                        response.addError("Failed to create issue for row " + rowNum + ": " + title);
                    }

                }
            }

            workbook.close();
        } catch (Exception e) {
            throw new IOException("Failed to process Excel file: " + e.getMessage(), e);
        }

        return createdIssueIds;
    }


}