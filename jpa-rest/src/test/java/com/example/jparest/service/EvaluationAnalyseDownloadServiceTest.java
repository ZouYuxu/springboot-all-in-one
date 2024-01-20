package com.example.jparest.service;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;
import org.slf4j.Logger;

import static org.mockito.Mockito.*;

class EvaluationAnalyseDownloadServiceTest {
    @Mock
    Logger log;
    @InjectMocks
    EvaluationAnalyseDownloadService evaluationAnalyseDownloadService;

    @BeforeEach
    void setUp() {
        MockitoAnnotations.openMocks(this);
    }

    @Test
    void testDownloadEvaluationAnalyse() throws Exception {
        evaluationAnalyseDownloadService.downloadEvaluationAnalyse();
    }

    @Test
    void testTemplate() throws Exception {
        evaluationAnalyseDownloadService.template();
    }
}

//Generated with love by TestMe :) Please report issues and submit feature requests at: http://weirddev.com/forum#!/testme