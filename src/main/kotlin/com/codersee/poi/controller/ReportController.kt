package com.codersee.poi.controller

import com.codersee.poi.service.ReportService
import org.springframework.http.HttpHeaders
import org.springframework.http.MediaType
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController

@RestController
@RequestMapping("/api/report")
class ReportController(
    private val reportService: ReportService
) {

    @PostMapping("/xlsx")
    fun generateXlsxReport(): ResponseEntity<ByteArray> {
        val report = reportService.generateXlsxReport()

        return createResponseEntity(report, "report.xlsx")
    }

    @PostMapping("/xls")
    fun generateXlsReport(): ResponseEntity<ByteArray> {
        val report = reportService.generateXlsReport()

        return createResponseEntity(report, "report.xls")
    }

    private fun createResponseEntity(
        report: ByteArray,
        fileName: String
    ): ResponseEntity<ByteArray> =
        ResponseEntity.ok()
            .contentType(MediaType.APPLICATION_OCTET_STREAM)
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"$fileName\"")
            .body(report)

}
