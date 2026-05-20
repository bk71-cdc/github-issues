package cdc.gov.github.controller;

import cdc.gov.github.dto.LoadTicketsRequest;
import cdc.gov.github.dto.LoadTicketsResponse;
import cdc.gov.github.dto.CreateIssuesRequest;
import cdc.gov.github.dto.UpdateTicketsRequest;
import cdc.gov.github.service.TicketService;
import cdc.gov.github.service.UpdateTicketService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api")
public class TicketController {

    @Autowired
    TicketService ticketService;

    @Autowired
    UpdateTicketService updateTicketService;

    @Autowired
    public TicketController(TicketService ticketService) {
        this.ticketService = ticketService;
    }

    @PostMapping("/loadTickets")
    public ResponseEntity<LoadTicketsResponse> loadTickets(@RequestBody LoadTicketsRequest request) {
        try {
            if (!request.isValid()) {
                LoadTicketsResponse errorResponse = new LoadTicketsResponse();
                errorResponse.setMessage("fileName is required");
                errorResponse.setTotalCount(0);
                errorResponse.setSuccessCount(0);
                errorResponse.setFailureCount(0);
                return ResponseEntity.badRequest().body(errorResponse);
            }
            
            LoadTicketsResponse response = ticketService.loadTicketsFromExcel(request.getFileName());
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            LoadTicketsResponse errorResponse = new LoadTicketsResponse();
            errorResponse.setMessage("Internal server error: " + e.getMessage());
            errorResponse.setTotalCount(0);
            errorResponse.setSuccessCount(0);
            errorResponse.setFailureCount(0);
            errorResponse.addError("Controller error: " + e.getMessage());
            return ResponseEntity.internalServerError().body(errorResponse);
        }
    }
    
    @PostMapping("/createIssueNo")
    public ResponseEntity<LoadTicketsResponse> createIssueNo(@RequestBody CreateIssuesRequest request) {
        try {
            if (!request.isValid()) {
                LoadTicketsResponse errorResponse = new LoadTicketsResponse();
                errorResponse.setMessage("numberOfIssues must be greater than 0");
                errorResponse.setTotalCount(0);
                errorResponse.setSuccessCount(0);
                errorResponse.setFailureCount(0);
                return ResponseEntity.badRequest().body(errorResponse);
            }
            
            LoadTicketsResponse response = ticketService.createEmptyIssues(request.getNumberOfIssues());
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            LoadTicketsResponse errorResponse = new LoadTicketsResponse();
            errorResponse.setMessage("Internal server error: " + e.getMessage());
            errorResponse.setTotalCount(0);
            errorResponse.setSuccessCount(0);
            errorResponse.setFailureCount(0);
            errorResponse.addError("Controller error: " + e.getMessage());
            return ResponseEntity.internalServerError().body(errorResponse);
        }
    }
    
    @PostMapping("/updateTickets")
    public ResponseEntity<LoadTicketsResponse> updateTickets(@RequestBody UpdateTicketsRequest request) {
        try {
            if (!request.isValid()) {
                LoadTicketsResponse errorResponse = new LoadTicketsResponse();
                errorResponse.setMessage("fileName is required");
                errorResponse.setTotalCount(0);
                errorResponse.setSuccessCount(0);
                errorResponse.setFailureCount(0);
                return ResponseEntity.badRequest().body(errorResponse);
            }
            
            LoadTicketsResponse response = updateTicketService.updateTicketsFromExcel(request.getFileName());
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            LoadTicketsResponse errorResponse = new LoadTicketsResponse();
            errorResponse.setMessage("Internal server error: " + e.getMessage());
            errorResponse.setTotalCount(0);
            errorResponse.setSuccessCount(0);
            errorResponse.setFailureCount(0);
            errorResponse.addError("Controller error: " + e.getMessage());
            return ResponseEntity.internalServerError().body(errorResponse);
        }
    }
}
