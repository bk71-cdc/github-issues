package cdc.gov.github.service;

import cdc.gov.github.dto.LoadTicketsResponse;

public interface UpdateTicketService {
    LoadTicketsResponse loadTicketsFromExcel(String fileName);
    LoadTicketsResponse createEmptyIssues(int numberOfIssues);
    LoadTicketsResponse updateTicketsFromExcel(String fileName);
}
