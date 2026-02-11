let allReqTable, pendingReqTable, returenedRequestTable;
let PSSC_DownloadRequest = `{{settings['Pssc_DownloadRequestExcel']}}`
let PSSC_DownloadApprovalRequest = `{{settings['Pssc_DownloadApprovalRequestExcel']}}`
let getAllNotificationUrl = `{{settings['PSSC_GetAllNotification']}}`
let markAsReadNotificaitonUrl = `{{settings['PSSC_MarkNotificationRead']}}`
let currentUserTeamType = null;
let grcApproveRequestList = [], grcRejectRequestList = [];
let getReturnedRequestList = [];
let grcApproveTable, grcRejectTable;
let p1Table;
let loggedInUserId;
// Pagination variables
let fullRequestList = []; // Store full request list
let loadedRequestCount = 0; // Track how many records have been loaded
const RECORDS_PER_BATCH = 50; // Number of records to load per batch
const RECORDS_PER_PAGE = 10; // Number of records to display per page in DataTable
let allPendingRequests = []; // Store all pending requests
let loadedPendingBatches = 0; // Track how many batches have been loaded to table
let allRequestDataStorage = null; // Store all request data for filtering (for all tabs)
let filteredRequestData = null; // Store filtered data after applying filters
let isLoadingBatch = false; // Flag to prevent duplicate batch loading
let lastCheckedPage = -1; // Track last page we checked to prevent duplicate checks
let batchLoadTimeout = null; // Debounce timer for batch loading

// Variables for other tabs (All, Resolved, Returned, Rejected, P1)
let allAllRequests = []; // Store all "All" tab requests
let loadedAllBatches = 0;
let isLoadingAllBatch = false;
let lastCheckedAllPage = -1;
let allAllBatchLoadTimeout = null;

let allResolvedRequests = []; // Store all "Resolved/Approved" tab requests
let loadedResolvedBatches = 0;
let isLoadingResolvedBatch = false;
let lastCheckedResolvedPage = -1;
let allResolvedBatchLoadTimeout = null;

let allReturnedRequests = []; // Store all "Returned" tab requests
let loadedReturnedBatches = 0;
let isLoadingReturnedBatch = false;
let lastCheckedReturnedPage = -1;
let allReturnedBatchLoadTimeout = null;

let allRejectedRequests = []; // Store all "Rejected" tab requests
let loadedRejectedBatches = 0;
let isLoadingRejectedBatch = false;
let lastCheckedRejectedPage = -1;
let allRejectedBatchLoadTimeout = null;

let allP1Requests = []; // Store all "P1" tab requests
let loadedP1Batches = 0;
let isLoadingP1Batch = false;
let lastCheckedP1Page = -1;
let allP1BatchLoadTimeout = null;
let userWithTeam = {};
let approvalRequest = [];
let UVP_portalurl = `{{ settings['uvpportalurl'] }}`;

// Simple date sorting function
function parseDateForSorting(data) {
    if (!data) return 0;

    // Try data-order attribute first (fastest)
    var match = data.match(/data-order="(\d+)"/);
    if (match) return parseInt(match[1]);

    // Convert DD/MM/YYYY HH:MM to proper date for sorting
    var dateStr = data.replace(/<[^>]*>/g, '').trim();
    var ddMMyyyy = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})(.*)/);
    if (ddMMyyyy) {
        var formattedDateStr = ddMMyyyy[3] + '-' + ddMMyyyy[2] + '-' + ddMMyyyy[1] + (ddMMyyyy[4] ? ddMMyyyy[4] : ' 00:00');
        return new Date(formattedDateStr).getTime();
    }

    return new Date(dateStr).getTime() || 0;
}

// Cache for SLA settings to avoid repeated API calls
let slaSettingsCache = null;

// Function to get SLA settings
async function getSLASettings() {
    if (slaSettingsCache) {
        return slaSettingsCache;
    }

    try {
        // Use the exact entity set name and field names from the working API call
        let slaSettings = await GetEntityList(
            "al_slasettingses", // Entity set name (plural with "es")
            "statecode eq 0", // Active SLA Settings
            "al_slasettingsid,al_name,al_slabreachminutes,al_subrequesttype" // Correct field names
        );
        // If GetEntityList returns null due to error, return empty array instead
        if (slaSettings === null) {
            console.warn("SLA settings fetch returned null, using empty array");
            slaSettingsCache = [];
            return [];
        }
        slaSettingsCache = slaSettings || [];
        return slaSettingsCache;
    } catch (error) {
        console.error("Error fetching SLA settings:", error);
        // Cache empty array to prevent repeated failed calls
        slaSettingsCache = [];
        return [];
    }
}

// Helper to format UVP Request ID cell with UVP portal link icon
function formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId) {
    if (!uvpRequestId || uvpRequestId === '-') {
        return uvpRequestId;
    }

    // Build dynamic UVP portal URL using vendor request id when available
    let baseUrl = UVP_portalurl.replace(/\/+$/, "") + "/vendor-approver";
    let uvpPortalUrl = baseUrl;

    if (uvpVendorRequestId) {
        uvpPortalUrl += `?vendorid=${encodeURIComponent(uvpVendorRequestId)}`;
    }

    // Small Font Awesome external-link icon next to the UVP Request ID
    return `<span class="uvpid-header">${uvpRequestId}</span>
        <a href="${uvpPortalUrl}" target="_blank" rel="noopener noreferrer" class="uvp-portal-link" title="Open in UVP portal" style="margin-left:4px; color:#67149d;">
            <i class="fas fa-external-link-alt" aria-hidden="true"></i>
        </a>`;
}

// Function to get assigned date/time for PSSC Requestor
function getAssignedDateTime(result) {
    try {
        // Priority 0: Latest approval request (handles send-back/reassign scenarios)
        if (result.al_approvalrequest_Request_al_request && result.al_approvalrequest_Request_al_request.length > 0) {
            let latestApproval = result.al_approvalrequest_Request_al_request
                .filter(ar => ar.createdon)
                .sort((a, b) => new Date(b.createdon) - new Date(a.createdon))[0];
            if (latestApproval) {
                let assignedDate = new Date(latestApproval.createdon);
                return {
                    date: assignedDate,
                    formatted: assignedDate.toLocaleString('en-US', {
                        year: 'numeric',
                        month: '2-digit',
                        day: '2-digit',
                        hour: '2-digit',
                        minute: '2-digit',
                        hour12: true
                    }),
                    timestamp: assignedDate.getTime()
                };
            }
        }

        // Priority 1: Check if there's an approval request with team type 3 (PSSC)
        if (result.al_approvalrequest_Request_al_request && result.al_approvalrequest_Request_al_request.length > 0) {
            // Find the first PSSC approval request (team type 3)
            let psscApprovalRequest = result.al_approvalrequest_Request_al_request.find(
                ar => ar.al_teamtype === 3 || ar.al_teamtype === "3"
            );
            if (psscApprovalRequest && psscApprovalRequest.createdon) {
                let assignedDate = new Date(psscApprovalRequest.createdon);
                return {
                    date: assignedDate,
                    formatted: assignedDate.toLocaleString('en-US', {
                        year: 'numeric',
                        month: '2-digit',
                        day: '2-digit',
                        hour: '2-digit',
                        minute: '2-digit',
                        hour12: true
                    }),
                    timestamp: assignedDate.getTime()
                };
            }

            // If no PSSC approval request found, check for any approval request with primary contact
            if (result._al_primarycontact_value) {
                let assignedApprovalRequest = result.al_approvalrequest_Request_al_request.find(
                    ar => ar._al_primarycontact_value === result._al_primarycontact_value
                );
                if (assignedApprovalRequest && assignedApprovalRequest.createdon) {
                    let assignedDate = new Date(assignedApprovalRequest.createdon);
                    return {
                        date: assignedDate,
                        formatted: assignedDate.toLocaleString('en-US', {
                            year: 'numeric',
                            month: '2-digit',
                            day: '2-digit',
                            hour: '2-digit',
                            minute: '2-digit',
                            hour12: true
                        }),
                        timestamp: assignedDate.getTime()
                    };
                }
            }
        }

        // Priority 2: If request has primary contact assigned, use modifiedon (when it was last assigned/modified)
        // This is a good indicator of when the request was assigned to the current assignee
        if (result._al_primarycontact_value && result.modifiedon) {
            let assignedDate = new Date(result.modifiedon);
            return {
                date: assignedDate,
                formatted: assignedDate.toLocaleString('en-US', {
                    year: 'numeric',
                    month: '2-digit',
                    day: '2-digit',
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: true
                }),
                timestamp: assignedDate.getTime()
            };
        }

        // Priority 3: Fallback to createdon if request is in assigned state
        if (result.al_requeststatus === 3 && result.createdon) {
            let assignedDate = new Date(result.createdon);
            return {
                date: assignedDate,
                formatted: assignedDate.toLocaleString('en-US', {
                    year: 'numeric',
                    month: '2-digit',
                    day: '2-digit',
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: true
                }),
                timestamp: assignedDate.getTime()
            };
        }

        // Priority 4: If request has any approval request, use the earliest one
        if (result.al_approvalrequest_Request_al_request && result.al_approvalrequest_Request_al_request.length > 0) {
            // Sort by createdon and get the earliest
            let sortedApprovals = result.al_approvalrequest_Request_al_request
                .filter(ar => ar.createdon)
                .sort((a, b) => new Date(a.createdon) - new Date(b.createdon));

            if (sortedApprovals.length > 0) {
                let assignedDate = new Date(sortedApprovals[0].createdon);
                return {
                    date: assignedDate,
                    formatted: assignedDate.toLocaleString('en-US', {
                        year: 'numeric',
                        month: '2-digit',
                        day: '2-digit',
                        hour: '2-digit',
                        minute: '2-digit',
                        hour12: true
                    }),
                    timestamp: assignedDate.getTime()
                };
            }
        }

        return {
            date: null,
            formatted: "-",
            timestamp: 0
        };
    } catch (error) {
        console.error("Error getting assigned date/time:", error);
        return {
            date: null,
            formatted: "-",
            timestamp: 0
        };
    }
}


// Function to calculate Due Date: Assigned Date & Time + SLA breach minutes
async function calculateDueDate(result) {
    try {
        if (!result.al_subrequesttype) {
            return {
                date: null,
                formatted: "-",
                timestamp: 0
            };
        }

        // Base time for SLA calculation: Use "Assigned Date & Time" (not Created On)
        // Due Date = Assigned Date & Time + SLA Breach Minutes
        let assignedDateTime = getAssignedDateTime(result);

        if (!assignedDateTime.date || !assignedDateTime.timestamp) {
            return {
                date: null,
                formatted: "-",
                timestamp: 0
            };
        }

        // Get SLA settings
        let slaSettings = await getSLASettings();
        if (!slaSettings || slaSettings.length === 0) {
            return {
                date: null,
                formatted: "-",
                timestamp: 0
            };
        }

        // Find SLA setting for this sub request type
        let slaSetting = slaSettings.find(
            s => s.al_subrequesttype === result.al_subrequesttype ||
                s.al_subrequesttype === String(result.al_subrequesttype)
        );

        // Get breach minutes from the correct field (al_slabreachminutes)
        let slaBreachMinutes = slaSetting ? slaSetting.al_slabreachminutes : null;

        if (!slaSetting || !slaBreachMinutes) {
            return {
                date: null,
                formatted: "-",
                timestamp: 0
            };
        }

        // Calculate Due Date: Assigned Date & Time + SLA Breach Minutes
        let dueDate = new Date(assignedDateTime.date.getTime() + (slaBreachMinutes * 60 * 1000));

        return {
            date: dueDate,
            formatted: dueDate.toLocaleString('en-US', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                hour12: true
            }),
            timestamp: dueDate.getTime()
        };
    } catch (error) {
        console.error("Error calculating due date:", error);
        return {
            date: null,
            formatted: "-",
            timestamp: 0
        };
    }
}

// Function to calculate hours remaining until SLA breach (kept for backward compatibility if needed)
async function calculateHoursRemainingUntilSLABreach(result, assignedDateTime) {
    try {
        if (!assignedDateTime.date || !result.al_subrequesttype) {
            return {
                hours: null,
                formatted: "-",
                sortValue: 0
            };
        }

        // Get SLA settings
        let slaSettings = await getSLASettings();
        if (!slaSettings || slaSettings.length === 0) {
            return {
                hours: null,
                formatted: "-",
                sortValue: 0
            };
        }

        // Find SLA setting for this sub request type
        let slaSetting = slaSettings.find(
            s => s.al_subrequesttype === result.al_subrequesttype ||
                s.al_subrequesttype === String(result.al_subrequesttype)
        );

        if (!slaSetting || !slaSetting.al_slabreachminutes) {
            return {
                hours: null,
                formatted: "-",
                sortValue: 0
            };
        }

        // Calculate SLA breach time (assigned date + SLA breach minutes)
        let slaBreachMinutes = slaSetting.al_slabreachminutes;
        let slaBreachDate = new Date(assignedDateTime.date.getTime() + (slaBreachMinutes * 60 * 1000));

        // Calculate hours remaining
        let now = new Date();
        let hoursRemaining = (slaBreachDate.getTime() - now.getTime()) / (1000 * 60 * 60);

        if (hoursRemaining < 0) {
            // SLA already breached
            return {
                hours: hoursRemaining,
                formatted: `<span class="text-danger">Breached (${Math.abs(hoursRemaining).toFixed(1)} hrs ago)</span>`,
                sortValue: hoursRemaining
            };
        } else {
            return {
                hours: hoursRemaining,
                formatted: `${hoursRemaining.toFixed(1)} hrs`,
                sortValue: hoursRemaining
            };
        }
    } catch (error) {
        console.error("Error calculating hours remaining:", error);
        return {
            hours: null,
            formatted: "-",
            sortValue: 0
        };
    }
}
$(async function () {
    allReqTable = new DataTable('#allReqTable', {
        order: [[3, 'desc']], // Sort by date column descending (newest first)
        autoWidth: false,
        pageLength: RECORDS_PER_PAGE, // Show 10 records per page
        paging: true, // Enable pagination
        pagingType: 'full_numbers', // Show full pagination controls
        dom: '<"top d-flex justify-content-between"fB>rt<"bottom d-flex justify-content-between"lip><"clear">', // Add button container (we'll hide them)
        buttons: [
            {
                extend: 'excelHtml5',
                text: 'Export', // Won't be visible, we'll trigger programmatically
                filename: function () {
                    let now = new Date();
                    let dateStr = now.toISOString().slice(0, 19).replace(/[:T]/g, '-');
                    return 'All_Requests_' + dateStr;
                },
                exportOptions: {
                    modifier: {
                        search: 'applied', // Only filtered rows
                        order: 'applied',  // Respect current sort
                        page: 'all'        // Ignore pagination, export all visible rows
                    },
                    columns: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14] // Only visible columns (including new columns after Requested By)
                }
            }
        ],
        columnDefs: [
            {
                targets: '_all',
                defaultContent: ''
            },
            {
                targets: 3, // Date column
                type: 'html-num-fmt',
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                // Vendor Name (4) and Legal Entity (7)
                targets: [4, 7],
                className: "ellipsis"
            },
            { targets: [15, 16, 17, 18, 19, 20, 21], visible: false } // Hide PR Number, PO Number, and other hidden columns
        ],
        data: [],
        drawCallback: function (settings) {
            const api = this.api();
            const pageInfo = api.page.info();
            let lastPageIndex = pageInfo.pages - 1; // zero-based
            const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1; // Zero-based, so add 1
            const end = pageInfo.end;
            const total = pageInfo.recordsDisplay; // After filtering

            // Check if filtered data exists, use that count; otherwise use original array length
            const totalRecords = (filteredRequestData && filteredRequestData.length > 0)
                ? filteredRequestData.length
                : (allAllRequests.length > 0
                    ? allAllRequests.length
                    : pageInfo.recordsTotal);

            // Update DataTables internal settings to reflect total records for pagination
            // This makes DataTables show pagination for all records, not just loaded ones
            if (allAllRequests.length > 0 || (filteredRequestData && filteredRequestData.length > 0)) {
                let dtSettings = api.settings()[0];
                if (dtSettings) {
                    // Force update the record counts - this is what DataTables uses for pagination
                    // Set these BEFORE DataTables recalculates
                    dtSettings._iRecordsTotal = totalRecords;
                    dtSettings._iRecordsDisplay = totalRecords;

                    // Also override the functions that DataTables uses to get record counts
                    if (!dtSettings._originalFnRecordsTotal) {
                        dtSettings._originalFnRecordsTotal = dtSettings.fnRecordsTotal;
                        dtSettings._originalFnRecordsDisplay = dtSettings.fnRecordsDisplay;
                    }
                    dtSettings.fnRecordsTotal = function () { return totalRecords; };
                    dtSettings.fnRecordsDisplay = function () { return totalRecords; };

                    let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                    if ($info.length) {
                        $info.text(`Showing ${start} to ${end} of ${totalRecords} entries`);
                    }

                    // Disable/enable last page button based on batch loading status
                    let $pagination = $(api.table().node()).parents('.table-responsive').find('.dt-paging');
                    let $lastButton = $pagination.find('button.dt-paging-button.last');
                    let $lastPageBtn = $pagination.find(
                        `button.dt-paging-button[data-dt-idx="${lastPageIndex}"]`
                    );
                    if ($lastButton.length) {
                        // Check if all batches are loaded
                        let totalBatchesNeeded = Math.ceil(totalRecords / RECORDS_PER_BATCH);
                        let allBatchesLoaded = loadedAllBatches >= totalBatchesNeeded || (loadedAllBatches * RECORDS_PER_BATCH >= totalRecords);

                        if (!allBatchesLoaded) {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .addClass('disabled')
                                    .attr('aria-disabled', 'true')
                                    .attr('tabindex', '-1')
                                    .prop('disabled', true)
                                    .css({
                                        'pointer-events': 'none',
                                        'opacity': '0.5'
                                    });
                            });
                        } else {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .removeClass('disabled')
                                    .attr('aria-disabled', 'false')
                                    .removeAttr('tabindex')
                                    .prop('disabled', false)
                                    .css({
                                        'pointer-events': '',
                                        'opacity': '1'
                                    });
                            });
                        }
                    }
                }

                // Check if we need to load next batch based on current page
                // Use debounce to prevent multiple rapid calls
                let currentPage = pageInfo.page + 1; // DataTables uses 0-based indexing
                if (!isLoadingAllBatch && currentPage !== lastCheckedAllPage) {
                    // Clear any pending timeout
                    if (allAllBatchLoadTimeout) {
                        clearTimeout(allAllBatchLoadTimeout);
                    }

                    let recordsPerPage = pageInfo.length;
                    let recordsNeeded = currentPage * recordsPerPage;
                    let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                    // Check if we need to load batches BEFORE debounce (to show loader immediately)
                    // Load if we need more batches AND we have more records than currently loaded
                    if (batchNeeded > loadedAllBatches && loadedAllBatches * RECORDS_PER_BATCH < allAllRequests.length && !isLoadingAllBatch) {
                        // Show loader IMMEDIATELY (synchronously, before any async operations)
                        let $loader = $('.custom-loader');
                        if ($loader.length) {
                            $loader.css("display", "flex");
                        }

                        // Hide "no records found" message immediately
                        let $table = $(api.table().node());
                        let $emptyRow = $table.find('tbody tr.dataTables_empty');
                        if ($emptyRow.length) {
                            $emptyRow.hide();
                        }
                    }

                    // Debounce the batch loading check
                    allAllBatchLoadTimeout = setTimeout(() => {
                        lastCheckedAllPage = currentPage;
                        let recordsPerPage = pageInfo.length;
                        let recordsNeeded = currentPage * recordsPerPage;
                        let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                        // Load all batches up to the needed one
                        // Load if we need more batches AND we have more records than currently loaded
                        if (batchNeeded > loadedAllBatches && loadedAllBatches * RECORDS_PER_BATCH < allAllRequests.length && !isLoadingAllBatch) {
                            // Save current page to restore after loading
                            let targetPage = currentPage - 1; // DataTables uses 0-based indexing

                            // Get loader reference (already shown above)
                            let $loader = $('.custom-loader');

                            // Hide "no records found" message during loading (in case it appeared)
                            let $table = $(api.table().node());
                            let $emptyRow = $table.find('tbody tr.dataTables_empty');
                            if ($emptyRow.length) {
                                $emptyRow.hide();
                            }

                            // Load batches asynchronously
                            (async () => {
                                try {
                                    while (loadedAllBatches < batchNeeded && !isLoadingAllBatch) {
                                        // Don't show loader for each individual batch, only show it once at the start
                                        if (typeof loadNextAllBatch === 'function') {
                                            await loadNextAllBatch(false);
                                            // The function updates loadedAllBatches after loadNextBatch completes
                                        } else {
                                            console.error('loadNextAllBatch is not accessible!');
                                            break;
                                        }
                                    }

                                    // After loading, ensure pagination is updated
                                    let dtSettings = api.settings()[0];
                                    if (dtSettings) {
                                        dtSettings._iRecordsTotal = allAllRequests.length;
                                        dtSettings._iRecordsDisplay = allAllRequests.length;
                                    }

                                    // Update last page button status after batch loading
                                    let $pagination = $(api.table().node()).parents('.table-responsive').find('.dt-paging');
                                    let $lastButton = $pagination.find('button.dt-paging-button.last');
                                    let $lastPageBtn = $pagination.find(
                                        `button.dt-paging-button[data-dt-idx="${lastPageIndex}"]`
                                    );
                                    if ($lastButton.length) {
                                        let totalBatchesNeeded = Math.ceil(allAllRequests.length / RECORDS_PER_BATCH);
                                        let allBatchesLoaded = loadedAllBatches >= totalBatchesNeeded || (loadedAllBatches * RECORDS_PER_BATCH >= allAllRequests.length);
                                        if (allBatchesLoaded) {
                                            $lastButton.removeClass('disabled')
                                                .attr('aria-disabled', 'false')
                                                .removeAttr('tabindex')
                                                .prop('disabled', false)
                                                .css('pointer-events', '')
                                                .css('opacity', '1');
                                            $lastPageBtn.removeClass('disabled')
                                                .attr('aria-disabled', 'false')
                                                .removeAttr('tabindex')
                                                .prop('disabled', false)
                                                .css('pointer-events', '')
                                                .css('opacity', '1');
                                        }
                                    }

                                    // Restore the page the user was on (use targetPage which is the page they clicked)
                                    // Use setTimeout to ensure data is fully appended before changing page
                                    setTimeout(() => {
                                        api.page(targetPage).draw('page');
                                    }, 100);
                                } finally {
                                    // Hide loader after everything is done
                                    if ($loader.length) {
                                        $loader.css("display", "none");
                                    }
                                }
                            })();
                        }
                    }, 100); // 100ms debounce
                }
            }

            // Update header text with total records
            $('h3.allReqCount').text(`Request ${start} - ${end} of ${totalRecords || total}`);
        },
        initComplete: function (settings, json) {
            // Set total records for pagination after initial load
            if (allReqTable && allAllRequests.length > 0) {
                try {
                    if (typeof allReqTable.api === 'function') {
                        let api = allReqTable.api();
                        if (api && api.settings && api.settings().length > 0) {
                            let dtSettings = api.settings()[0];
                            if (dtSettings) {
                                dtSettings._iRecordsTotal = allAllRequests.length;
                                dtSettings._iRecordsDisplay = allAllRequests.length;
                                // Override DataTables' internal functions to always return the total count
                                dtSettings.fnRecordsTotal = function () { return allAllRequests.length; };
                                dtSettings.fnRecordsDisplay = function () { return allAllRequests.length; };
                                api.draw(false); // Redraw to update pagination controls
                            }
                        }
                    } else {
                        // If not a DataTable instance, try to get it via jQuery
                        let $table = $('#allReqTable');
                        if ($table.length && $.fn.DataTable.isDataTable('#allReqTable')) {
                            let api = $table.DataTable();
                            let dtSettings = api.settings()[0];
                            if (dtSettings) {
                                dtSettings._iRecordsTotal = allAllRequests.length;
                                dtSettings._iRecordsDisplay = allAllRequests.length;
                                // Override DataTables' internal functions to always return the total count
                                dtSettings.fnRecordsTotal = function () { return allAllRequests.length; };
                                dtSettings.fnRecordsDisplay = function () { return allAllRequests.length; };
                                api.draw(false);
                            }
                        }
                    }
                } catch (error) {
                    console.warn('Error updating pagination settings after initial load for All tab:', error);
                }
            }
        },
        infoCallback: function (settings, start, end, max, total, pre) {
            // Check if filtered data exists, use that count; otherwise use original array length
            let displayTotal = (filteredRequestData && filteredRequestData.length > 0)
                ? filteredRequestData.length
                : (allAllRequests.length > 0 ? allAllRequests.length : total);
            let startNum = start + 1;
            let endNum = Math.min(end, displayTotal);
            return `Showing ${startNum} to ${endNum} of ${displayTotal} entries`;
        },
        createdRow: function (row, data, dataIndex) {
            const raw = data.__raw;
            if (!raw) return; // safety check
            $('td', row).each(function (index) {
                // Add title for ellipsis columns (Vendor Name at 4 and Legal Entity at 7)
                if (index === 4 || index === 7) {
                    const cellText = $(this).text().trim();
                    if (cellText) {
                        $(this).attr('title', cellText);
                    }
                }
                if (index >= 1 && index <= 10) {
                    const vendorname = raw.al_vendormultiplelabel || raw.al_vendorname || raw.al_eauctionname;
                    if (raw.al_subrequesttype === 12) {
                        const tooltipText = `PO Number: ${raw.al_ponumber || 'N/A'}\nPR Number: ${raw.al_prnumber || 'N/A'}\nVendor Name: ${vendorname || 'N/A'}\nProject Id: ${raw.al_projectnameid || 'N/A'}`;
                        $(this).addClass('custom-row-tooltip');
                        $(this).attr('data-tooltip', tooltipText);
                    }
                }
            });
        }
    });

    // Common batch loading function - reduces code duplication
    async function loadNextBatch(config) {
        const {
            allRecords,
            loadedBatches,
            isLoadingFlag,
            tableSelector,
            tableInstance,
            bindRequestTeamType = teamType,
            bindRequestSkip = null,
            bindRequestTop = null
        } = config;

        if (allRecords.length === 0 || isLoadingFlag.value) return;
        isLoadingFlag.value = true;

        let $loader = null;
        if (config.showLoader) {
            $loader = $('.custom-loader');
            if ($loader.length) {
                $loader.css("display", "flex");
            }
        }

        let $table = $(tableSelector);
        let $emptyRow = $table.find('tbody tr.dataTables_empty');
        if ($emptyRow.length) {
            $emptyRow.hide();
        }

        try {
            let startIndex = loadedBatches.value * RECORDS_PER_BATCH;
            let endIndex = startIndex + RECORDS_PER_BATCH;
            let nextBatch = allRecords.slice(startIndex, endIndex);

            if (nextBatch.length > 0) {
                await BindRequest(userWithTeam, nextBatch, tableInstance, bindRequestTeamType, bindRequestSkip, bindRequestTop, true);
                loadedBatches.value++;
            }
        } finally {
            isLoadingFlag.value = false;
            if ($loader && $loader.length) {
                $loader.css("display", "none");
            }
        }
    }

    // Wrapper functions for each tab (using common function)
    async function loadNextPendingBatch(showLoader = false) {
        const config = {
            allRecords: allPendingRequests,
            loadedBatches: { value: loadedPendingBatches },
            isLoadingFlag: { value: isLoadingBatch },
            tableSelector: '#pendingReqTable',
            tableInstance: pendingReqTable,
            bindRequestTeamType: -1,
            showLoader
        };
        await loadNextBatch(config);
        loadedPendingBatches = config.loadedBatches.value;
        isLoadingBatch = config.isLoadingFlag.value;
    }

    async function loadNextAllBatch(showLoader = false) {
        const config = {
            allRecords: allAllRequests,
            loadedBatches: { value: loadedAllBatches },
            isLoadingFlag: { value: isLoadingAllBatch },
            tableSelector: '#allReqTable',
            tableInstance: allReqTable,
            bindRequestTeamType: -1,
            showLoader
        };
        await loadNextBatch(config);
        loadedAllBatches = config.loadedBatches.value;
        isLoadingAllBatch = config.isLoadingFlag.value;
    }

    async function loadNextResolvedBatch(showLoader = false) {
        const config = {
            allRecords: allResolvedRequests,
            loadedBatches: { value: loadedResolvedBatches },
            isLoadingFlag: { value: isLoadingResolvedBatch },
            tableSelector: '#grcApproveTable',
            tableInstance: grcApproveTable,
            bindRequestTeamType: -1,
            showLoader
        };
        await loadNextBatch(config);
        loadedResolvedBatches = config.loadedBatches.value;
        isLoadingResolvedBatch = config.isLoadingFlag.value;
    }

    async function loadNextReturnedBatch(showLoader = false) {
        const config = {
            allRecords: allReturnedRequests,
            loadedBatches: { value: loadedReturnedBatches },
            isLoadingFlag: { value: isLoadingReturnedBatch },
            tableSelector: '#returenedRequestTable',
            tableInstance: returenedRequestTable,
            bindRequestTeamType: -1,
            showLoader
        };
        await loadNextBatch(config);
        loadedReturnedBatches = config.loadedBatches.value;
        isLoadingReturnedBatch = config.isLoadingFlag.value;
    }

    async function loadNextRejectedBatch(showLoader = false) {
        const config = {
            allRecords: allRejectedRequests,
            loadedBatches: { value: loadedRejectedBatches },
            isLoadingFlag: { value: isLoadingRejectedBatch },
            tableSelector: '#grcRejectTable',
            tableInstance: grcRejectTable,
            bindRequestTeamType: -1,
            showLoader
        };
        await loadNextBatch(config);
        loadedRejectedBatches = config.loadedBatches.value;
        isLoadingRejectedBatch = config.isLoadingFlag.value;
    }

    async function loadNextP1Batch(showLoader = false) {
        const config = {
            allRecords: allP1Requests,
            loadedBatches: { value: loadedP1Batches },
            isLoadingFlag: { value: isLoadingP1Batch },
            tableSelector: '#p1Table',
            tableInstance: p1Table,
            bindRequestTeamType: -1,
            showLoader
        };
        await loadNextBatch(config);
        loadedP1Batches = config.loadedBatches.value;
        isLoadingP1Batch = config.isLoadingFlag.value;
    }

    // Common function to update DataTables pagination settings
    function updateDataTablesPagination(tableInstance, totalRecords) {
        if (!tableInstance || totalRecords === undefined) return;
        try {
            if (typeof tableInstance.api === 'function') {
                let api = tableInstance.api();
                if (api && api.settings && api.settings().length > 0) {
                    let dtSettings = api.settings()[0];
                    if (dtSettings) {
                        dtSettings._iRecordsTotal = totalRecords;
                        dtSettings._iRecordsDisplay = totalRecords;
                        if (!dtSettings._originalFnRecordsTotal) {
                            dtSettings._originalFnRecordsTotal = dtSettings.fnRecordsTotal;
                            dtSettings._originalFnRecordsDisplay = dtSettings.fnRecordsDisplay;
                        }
                        dtSettings.fnRecordsTotal = function () { return totalRecords; };
                        dtSettings.fnRecordsDisplay = function () { return totalRecords; };
                        api.draw(false);
                    }
                }
            }
        } catch (error) {
            console.warn('Error updating pagination settings:', error);
        }
    }

    // Make function globally accessible
    window.updateDataTablesPagination = updateDataTablesPagination;

    // Common function to load first batch and update pagination
    async function loadFirstBatchAndUpdatePagination(config) {
        const {
            allRecords,
            loadedBatches,
            tableInstance,
            bindRequestTeamType = -1,
            bindRequestSkip = null,
            bindRequestTop = null
        } = config;

        loadedBatches.value = 0;
        let firstBatch = allRecords.slice(0, RECORDS_PER_BATCH);
        await BindRequest(userWithTeam, firstBatch, tableInstance, bindRequestTeamType, bindRequestSkip, bindRequestTop);
        loadedBatches.value = 1;
        window.updateDataTablesPagination(tableInstance, allRecords.length);
    }

    // Make function globally accessible
    window.loadFirstBatchAndUpdatePagination = loadFirstBatchAndUpdatePagination;

    // Common drawCallback factory function
    function createBatchLoadingDrawCallback(config) {
        const {
            allRecords,
            loadedBatches,
            isLoadingFlag,
            lastCheckedPage,
            batchLoadTimeout,
            loadNextBatchFn,
            headerSelector
        } = config;

        return function (settings) {
            const api = this.api();
            const pageInfo = api.page.info();
            let lastPageIndex = pageInfo.pages - 1; // zero-based
            const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1;
            const end = pageInfo.end;
            const total = pageInfo.recordsDisplay;

            const totalRecords = allRecords.length > 0 ? allRecords.length : pageInfo.recordsTotal;

            // Update DataTables internal settings
            if (allRecords.length > 0) {
                let dtSettings = api.settings()[0];
                if (dtSettings) {
                    dtSettings._iRecordsTotal = totalRecords;
                    dtSettings._iRecordsDisplay = totalRecords;
                    if (!dtSettings._originalFnRecordsTotal) {
                        dtSettings._originalFnRecordsTotal = dtSettings.fnRecordsTotal;
                        dtSettings._originalFnRecordsDisplay = dtSettings.fnRecordsDisplay;
                    }
                    dtSettings.fnRecordsTotal = function () { return totalRecords; };
                    dtSettings.fnRecordsDisplay = function () { return totalRecords; };
                    let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                    if ($info.length) {
                        $info.text(`Showing ${start} to ${end} of ${totalRecords} entries`);
                    }
                }

                // Check if we need to load next batch
                let currentPage = pageInfo.page + 1;
                if (!isLoadingFlag.value && currentPage !== lastCheckedPage.value) {
                    if (batchLoadTimeout.value) {
                        clearTimeout(batchLoadTimeout.value);
                    }
                    let recordsPerPage = pageInfo.length;
                    let recordsNeeded = currentPage * recordsPerPage;
                    let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                    if (batchNeeded > loadedBatches.value && batchNeeded * RECORDS_PER_BATCH <= allRecords.length && !isLoadingFlag.value) {
                        let $loader = $('.custom-loader');
                        if ($loader.length) {
                            $loader.css("display", "flex");
                        }
                        let $table = $(api.table().node());
                        let $emptyRow = $table.find('tbody tr.dataTables_empty');
                        if ($emptyRow.length) {
                            $emptyRow.hide();
                        }
                    }

                    batchLoadTimeout.value = setTimeout(() => {
                        lastCheckedPage.value = currentPage;
                        let recordsPerPage = pageInfo.length;
                        let recordsNeeded = currentPage * recordsPerPage;
                        let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                        if (batchNeeded > loadedBatches.value && batchNeeded * RECORDS_PER_BATCH <= allRecords.length && !isLoadingFlag.value) {
                            let targetPage = currentPage - 1;
                            let $loader = $('.custom-loader');
                            let $table = $(api.table().node());
                            let $emptyRow = $table.find('tbody tr.dataTables_empty');
                            if ($emptyRow.length) {
                                $emptyRow.hide();
                            }

                            (async () => {
                                try {
                                    while (loadedBatches.value < batchNeeded && !isLoadingFlag.value) {
                                        await loadNextBatchFn(false);
                                    }
                                    let dtSettings = api.settings()[0];
                                    if (dtSettings) {
                                        dtSettings._iRecordsTotal = allRecords.length;
                                        dtSettings._iRecordsDisplay = allRecords.length;
                                    }
                                    setTimeout(() => {
                                        api.page(targetPage).draw('page');
                                    }, 100);
                                } finally {
                                    if ($loader.length) {
                                        $loader.css("display", "none");
                                    }
                                }
                            })();
                        }
                    }, 100);
                }
            }

            // Update header text
            if (headerSelector) {
                $(headerSelector).text(`Request ${start} - ${end} of ${totalRecords || total}`);
            }
        };
    }

    pendingReqTable = new DataTable('#pendingReqTable', {
        order: [[4, 'desc']], // Sort by date column descending (newest first)
        data: [],
        autoWidth: false,
        pageLength: RECORDS_PER_PAGE, // Show 10 records per page
        paging: true, // Enable pagination
        pagingType: 'full_numbers', // Show full pagination controls
        dom: '<"top d-flex justify-content-between"fB>rt<"bottom d-flex justify-content-between"lip><"clear">', // Add button container (we'll hide them)
        buttons: [
            {
                extend: 'excelHtml5',
                text: 'Export', // Won't be visible, we'll trigger programmatically
                filename: function () {
                    let now = new Date();
                    let dateStr = now.toISOString().slice(0, 19).replace(/[:T]/g, '-');
                    return 'Pending_Requests_' + dateStr;
                },
                exportOptions: {
                    modifier: {
                        search: 'applied', // Only filtered rows
                        order: 'applied',  // Respect current sort
                        page: 'all'        // Ignore pagination, export all visible rows
                    },
                    columns: [1, 2, 3, 4, 5, 6, 7, 8, 9, 13, 14, 15] // Exclude checkbox (0), Requested By (10), Assigned Date (11), Due Date (12); Include: Request ID, UVP Request ID, Request Sub Type, Date & Time, Vendor Name, Vendor Code, Vendor Email, Legal Entity, Assigned To, Approval Status, Final Status, Action
                }
            }
        ],
        initComplete: function (settings, json) {
            // Force show Approval Status, Final Status, and Action headers immediately after initialization
            // This runs before any data is loaded, ensuring headers are visible from the start
            let $table = $('#pendingReqTable');
            let $headerRow = $table.find('thead tr');
            if ($headerRow.length > 0) {
                let $th10 = $headerRow.find('th').eq(10); // Approval Status
                let $th11 = $headerRow.find('th').eq(11); // Final Status
                let $th12 = $headerRow.find('th').eq(12); // Action

                if ($th10.length) {
                    $th10.show();
                    $th10.css('display', '');
                    $th10.css('visibility', 'visible');
                    $th10.removeClass('d-none');
                    $th10.css('width', 'auto');
                }
                if ($th11.length) {
                    $th11.show();
                    $th11.css('display', '');
                    $th11.css('visibility', 'visible');
                    $th11.removeClass('d-none');
                    $th11.css('width', 'auto');
                }
                if ($th12.length) {
                    $th12.show();
                    $th12.css('display', '');
                    $th12.css('visibility', 'visible');
                    $th12.removeClass('d-none');
                    $th12.css('width', 'auto');
                }
            }
        },
        infoCallback: function (settings, start, end, max, total, pre) {
            // Check if filtered data exists, use that count; otherwise use original array length
            let displayTotal = (filteredRequestData && filteredRequestData.length > 0)
                ? filteredRequestData.length
                : (allPendingRequests.length > 0 ? allPendingRequests.length : total);
            let startNum = start + 1;
            let endNum = Math.min(end, displayTotal);
            return `Showing ${startNum} to ${endNum} of ${displayTotal} entries`;

        },
        // drawCallback: function (settings) {
        //     // Update pagination info to show total records
        //     if (allPendingRequests.length > 0) {
        //         let api = this.api();
        //         let pageInfo = api.page.info();
        //         let start = pageInfo.start + 1;
        //         let end = Math.min(pageInfo.end, allPendingRequests.length);
        //         let total = allPendingRequests.length;

        //         // Update the info text manually
        //         let infoElement = $(this.api().table().node()).closest('.dataTables_wrapper').find('.dataTables_info');
        //         if (infoElement.length) {
        //             infoElement.text(`Showing ${start} to ${end} of ${total} entries`);
        //         }

        //         // Update header text
        //         $('h3.pendingReqCount').text(`Request ${start} - ${end} of ${total}`);
        //     }

        //     // Existing drawCallback logic
        //     const pageInfo = this.api().page.info();
        //     const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1;
        //     const end = pageInfo.end;
        //     const total = allPendingRequests.length > 0 ? allPendingRequests.length : pageInfo.recordsDisplay;

        //     // Ensure Assigned To (column 9) is visible in pending table
        //     if (this.api().table().node().id === 'pendingReqTable') {
        //         this.api().column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        //     }

        //     $('#pendingReqTable tbody td').each(function () {
        //         const $cell = $(this);
        //         const text = $cell.text().trim();
        //         // Check for truncation
        //         if (this.offsetWidth < this.scrollWidth) {
        //             $cell.attr('title', text);
        //         }
        //     });
        // },
        columnDefs: [
            {
                targets: '_all',
                defaultContent: ''
            },
            {
                targets: 0,
                width: "75px", // You can set it to "20%", "10rem", etc.
                className: "min-width-col"
            },
            {
                targets: 1,
                width: "80px" // You can set it to "20%", "10rem", etc.
            },
            {
                targets: 2,
                width: "140px" // You can set it to "20%", "10rem", etc.
            },
            {
                targets: 4, // Date column
                type: 'html-num-fmt',
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: [5, 6, 7, 8], // Vendor Name (5), Vendor Code (6), Vendor Email (7), and Legal Entity (8)
                className: "ellipsis"
            },
            {
                targets: [9], // Assigned To (9) - ensure this is visible
                visible: true,
                className: '' // Remove any classes that might hide it
            },
            {
                targets: [11, 12], // Assigned Date (11), Due Date (12) - conditionally visible for PSSC teams
                visible: false, // Default to hidden, will be shown for PSSC teams in BindRequest
                orderable: true,
                searchable: true
            },
            {
                targets: [13, 14, 15], // Approval Status (13), Final Status (14), Action (15) - ensure these are visible
                visible: true,
                className: '' // Remove any classes that might hide them
            },
            { targets: [0, 16, 17, 18, 19, 20], visible: false } // Hide checkbox (0), PR Number (16), PO Number (17), and other hidden columns
        ],
        drawCallback: function (settings) {
            const api = this.api();
            const pageInfo = api.page.info();
            const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1; // Zero-based, so add 1
            const end = pageInfo.end;
            const total = pageInfo.recordsDisplay; // After filtering
            let lastPageIndex = pageInfo.pages - 1; // zero-based
            // Check if filtered data exists, use that count; otherwise use original array length
            const totalRecords = (filteredRequestData && filteredRequestData.length > 0)
                ? filteredRequestData.length
                : (allPendingRequests.length > 0
                    ? allPendingRequests.length
                    : pageInfo.recordsTotal);

            // Update DataTables internal settings to reflect total records for pagination
            // This makes DataTables show pagination for all records, not just loaded ones
            if (allPendingRequests.length > 0 || (filteredRequestData && filteredRequestData.length > 0)) {
                let dtSettings = api.settings()[0];
                if (dtSettings) {
                    // Force update the record counts - this is what DataTables uses for pagination
                    // Set these BEFORE DataTables recalculates
                    dtSettings._iRecordsTotal = totalRecords;
                    dtSettings._iRecordsDisplay = totalRecords;

                    // Also override the functions that DataTables uses to get record counts
                    if (!dtSettings._originalFnRecordsTotal) {
                        dtSettings._originalFnRecordsTotal = dtSettings.fnRecordsTotal;
                        dtSettings._originalFnRecordsDisplay = dtSettings.fnRecordsDisplay;
                    }
                    dtSettings.fnRecordsTotal = function () { return totalRecords; };
                    dtSettings.fnRecordsDisplay = function () { return totalRecords; };

                    // Update the pagination info display manually
                    let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                    if ($info.length) {
                        $info.text(`Showing ${start} to ${end} of ${totalRecords} entries`);
                    }

                    // Disable/enable last page button based on batch loading status
                    let $pagination = $(api.table().node()).parents('.table-responsive').find('.dt-paging');
                    let $lastButton = $pagination.find('button.dt-paging-button.last');
                    let $lastPageBtn = $pagination.find(
                        `button.dt-paging-button[data-dt-idx="${lastPageIndex}"]`
                    );
                    if ($lastButton.length) {
                        // Check if all batches are loaded
                        let totalBatchesNeeded = Math.ceil(totalRecords / RECORDS_PER_BATCH);
                        let allBatchesLoaded = loadedPendingBatches >= totalBatchesNeeded || (loadedPendingBatches * RECORDS_PER_BATCH >= totalRecords);

                        if (!allBatchesLoaded) {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .addClass('disabled')
                                    .attr('aria-disabled', 'true')
                                    .attr('tabindex', '-1')
                                    .prop('disabled', true)
                                    .css({
                                        'pointer-events': 'none',
                                        'opacity': '0.5'
                                    });
                            });
                        } else {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .removeClass('disabled')
                                    .attr('aria-disabled', 'false')
                                    .removeAttr('tabindex')
                                    .prop('disabled', false)
                                    .css({
                                        'pointer-events': '',
                                        'opacity': '1'
                                    });
                            });
                        }
                    }
                }

                // Check if we need to load next batch based on current page
                // Use debounce to prevent multiple rapid calls
                let currentPage = pageInfo.page + 1; // DataTables uses 0-based indexing
                if (!isLoadingBatch && currentPage !== lastCheckedPage) {
                    // Clear any pending timeout
                    if (batchLoadTimeout) {
                        clearTimeout(batchLoadTimeout);
                    }

                    let recordsPerPage = pageInfo.length;
                    let recordsNeeded = currentPage * recordsPerPage;
                    let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                    // Check if we need to load batches BEFORE debounce (to show loader immediately)
                    // Load if we need more batches AND we have more records than currently loaded
                    if (batchNeeded > loadedPendingBatches && loadedPendingBatches * RECORDS_PER_BATCH < allPendingRequests.length && !isLoadingBatch) {
                        // Show loader IMMEDIATELY (synchronously, before any async operations)
                        let $loader = $('.custom-loader');
                        if ($loader.length) {
                            $loader.css("display", "flex");
                        }

                        // Hide "no records found" message immediately
                        let $table = $(api.table().node());
                        let $emptyRow = $table.find('tbody tr.dataTables_empty');
                        if ($emptyRow.length) {
                            $emptyRow.hide();
                        }
                    }

                    // Debounce the batch loading check
                    batchLoadTimeout = setTimeout(() => {
                        lastCheckedPage = currentPage;
                        let recordsPerPage = pageInfo.length;
                        let recordsNeeded = currentPage * recordsPerPage;
                        let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                        // Load all batches up to the needed one
                        // Load if we need more batches AND we have more records than currently loaded
                        if (batchNeeded > loadedPendingBatches && loadedPendingBatches * RECORDS_PER_BATCH < allPendingRequests.length && !isLoadingBatch) {
                            // Save current page to restore after loading
                            let targetPage = currentPage - 1; // DataTables uses 0-based indexing

                            // Get loader reference (already shown above)
                            let $loader = $('.custom-loader');

                            // Hide "no records found" message during loading (in case it appeared)
                            let $table = $(api.table().node());
                            let $emptyRow = $table.find('tbody tr.dataTables_empty');
                            if ($emptyRow.length) {
                                $emptyRow.hide();
                            }

                            // Load batches asynchronously
                            (async () => {
                                try {
                                    while (loadedPendingBatches < batchNeeded && !isLoadingBatch) {
                                        // Don't show loader for each individual batch, only show it once at the start
                                        await loadNextPendingBatch(false);
                                    }

                                    // After loading, ensure pagination is updated
                                    let dtSettings = api.settings()[0];
                                    if (dtSettings) {
                                        dtSettings._iRecordsTotal = allPendingRequests.length;
                                        dtSettings._iRecordsDisplay = allPendingRequests.length;
                                    }

                                    // Update last page button status after batch loading
                                    let $pagination = $(api.table().node()).parents('.table-responsive').find('.dt-paging');
                                    let $lastButton = $pagination.find('button.dt-paging-button.last');
                                    let $lastPageBtn = $pagination.find(
                                        `button.dt-paging-button[data-dt-idx="${lastPageIndex}"]`
                                    );
                                    if ($lastButton.length) {
                                        let totalBatchesNeeded = Math.ceil(allPendingRequests.length / RECORDS_PER_BATCH);
                                        let allBatchesLoaded = loadedPendingBatches >= totalBatchesNeeded || (loadedPendingBatches * RECORDS_PER_BATCH >= allPendingRequests.length);
                                        if (allBatchesLoaded) {
                                            $lastButton.removeClass('disabled')
                                                .attr('aria-disabled', 'false')
                                                .removeAttr('tabindex')
                                                .prop('disabled', false)
                                                .css('pointer-events', '')
                                                .css('opacity', '1');
                                            $lastPageBtn.removeClass('disabled')
                                                .attr('aria-disabled', 'false')
                                                .removeAttr('tabindex')
                                                .prop('disabled', false)
                                                .css('pointer-events', '')
                                                .css('opacity', '1');
                                        }
                                    }

                                    // Restore the page the user was on (use targetPage which is the page they clicked)
                                    // Use setTimeout to ensure data is fully appended before changing page
                                    setTimeout(() => {
                                        api.page(targetPage).draw('page');
                                    }, 100);
                                } finally {
                                    // Hide loader after everything is done
                                    if ($loader.length) {
                                        $loader.css("display", "none");
                                    }
                                }
                            })();
                        }
                    }, 100); // 100ms debounce
                }
            }

            const recordsPerPage = pageInfo.length;
            const pagingText = `${totalRecords} / ${recordsPerPage}`;

            // Put it wherever you want (example: next to pagination)
            let $wrapper = $(api.table().node()).parents('.table-responsive');

            let $customPaging = $wrapper.find('.customPagingInfo');
            if (!$customPaging.length) {
                $customPaging = $('<div class="customPagingInfo fw-bold ms-2"></div>');
                $wrapper.find('.dataTables_paginate').append($customPaging);
            }

            $customPaging.text(pagingText);

            // Update header text with total records
            $('h3.pendingReqCount').text(`Request ${start} - ${end} of ${totalRecords}`);

            // Ensure Assigned To (column 9) is visible in pending table
            if (this.api().table().node().id === 'pendingReqTable') {
                this.api().column(9).visible(true); // Assigned To - ALWAYS VISIBLE
            }

            $('#pendingReqTable tbody td').each(function () {
                const $cell = $(this);
                const text = $cell.text().trim();
                // Check for truncation
                if (this.offsetWidth < this.scrollWidth) {
                    $cell.attr('title', text); // native tooltip
                } else {
                    $cell.removeAttr('title'); // avoid redundant tooltips
                }
            });
        },
        createdRow: function (row, data, dataIndex) {
            const raw = data.__raw;
            if (!raw) return; // safety check
            const $cells = $('td', row);
            $cells.each(function (index) {
                // Add title for ellipsis columns (Vendor Name, Vendor Code, Vendor Email, and Legal Entity)
                if (index === 5 || index === 6 || index === 7 || index === 8) {
                    const cellText = $(this).text().trim();
                    if (cellText) {
                        $(this).attr('title', cellText);
                    }
                }
                if (index >= 2 && index <= 8 && index < $cells.length - 1) {
                    const vendorname = raw.al_vendormultiplelabel || raw.al_vendorname || raw.al_eauctionname;
                    if (raw.al_subrequesttype === 12) {
                        const tooltipText = `PO Number: ${raw.al_ponumber || 'N/A'}\nPR Number: ${raw.al_prnumber || 'N/A'}\nVendor Name: ${vendorname || 'N/A'}\nProject Id: ${raw.al_projectnameid || 'N/A'}`;
                        // Get existing cell text
                        const cellContent = $(this).html();
                        // Replace content with wrapped div
                        $(this).html(`<div class="custom-row-tooltip" data-tooltip="${tooltipText}">${cellContent}</div>`);
                    }
                }
            });
        }
    });

    // Force show Approval Status, Final Status, and Action headers immediately after table initialization
    // This ensures headers are visible even before data loads
    setTimeout(function () {
        let $pendingTable = $('#pendingReqTable');
        let $headerRow = $pendingTable.find('thead tr');
        if ($headerRow.length > 0) {
            let $th10 = $headerRow.find('th').eq(10); // Approval Status
            let $th11 = $headerRow.find('th').eq(11); // Final Status
            let $th12 = $headerRow.find('th').eq(12); // Action

            if ($th10.length) {
                $th10.show();
                $th10.css('display', '');
                $th10.css('visibility', 'visible');
                $th10.removeClass('d-none');
            }
            if ($th11.length) {
                $th11.show();
                $th11.css('display', '');
                $th11.css('visibility', 'visible');
                $th11.removeClass('d-none');
            }
            if ($th12.length) {
                $th12.show();
                $th12.css('display', '');
                $th12.css('visibility', 'visible');
                $th12.removeClass('d-none');
            }
        }
    }, 100);

    grcApproveTable = new DataTable('#grcApproveTable', {
        order: [[3, 'desc']], // Sort by date column descending (newest first)
        autoWidth: false,
        pageLength: RECORDS_PER_PAGE, // Show 10 records per page
        paging: true, // Enable pagination
        pagingType: 'full_numbers', // Show full pagination controls
        dom: '<"top d-flex justify-content-between"fB>rt<"bottom d-flex justify-content-between"lip><"clear">', // Add button container (we'll hide them)
        buttons: [
            {
                extend: 'excelHtml5',
                text: 'Export', // Won't be visible, we'll trigger programmatically
                filename: function () {
                    let now = new Date();
                    let dateStr = now.toISOString().slice(0, 19).replace(/[:T]/g, '-');
                    return 'Resolved_Requests_' + dateStr;
                },
                exportOptions: {
                    modifier: {
                        search: 'applied', // Only filtered rows
                        order: 'applied',  // Respect current sort
                        page: 'all'        // Ignore pagination, export all visible rows
                    },
                    columns: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] // Only visible columns (including new columns after Requested By)
                }
            }
        ],
        columnDefs: [
            {
                targets: '_all',
                defaultContent: ''
            },
            {
                targets: 0, // Legal Entity column index
                width: "50px", // You can set it to "20%", "10rem", etc.
                className: "min-width-col"
            },
            {
                targets: 3, // Date column
                type: 'html-num-fmt',
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                // Vendor Name (4), Vendor Code (5), Vendor Email (6)
                targets: [4, 5, 6],
                className: "ellipsis"
            },
            {
                targets: 9, // Requested By column - hide same as pending tab
                visible: false
            },
            {
                targets: 10, // Assigned Date & Time column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: 11, // Due Date column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            { targets: [15, 16, 17, 18, 19, 20, 21], visible: false } // Hide PR Number, PO Number, and other hidden columns
        ],
        drawCallback: function (settings) {
            const api = this.api();
            const pageInfo = api.page.info();
            let lastPageIndex = pageInfo.pages - 1; // zero-based
            const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1; // Zero-based, so add 1
            const end = pageInfo.end;
            const total = pageInfo.recordsDisplay; // After filtering

            // Check if filtered data exists, use that count; otherwise use original array length
            const totalRecords = (filteredRequestData && filteredRequestData.length > 0)
                ? filteredRequestData.length
                : (allResolvedRequests.length > 0
                    ? allResolvedRequests.length
                    : pageInfo.recordsTotal);

            // Update DataTables internal settings to reflect total records for pagination
            // This makes DataTables show pagination for all records, not just loaded ones
            if (allResolvedRequests.length > 0 || (filteredRequestData && filteredRequestData.length > 0)) {
                let dtSettings = api.settings()[0];
                if (dtSettings) {
                    // Force update the record counts - this is what DataTables uses for pagination
                    // Set these BEFORE DataTables recalculates
                    dtSettings._iRecordsTotal = totalRecords;
                    dtSettings._iRecordsDisplay = totalRecords;

                    // Also override the functions that DataTables uses to get record counts
                    if (!dtSettings._originalFnRecordsTotal) {
                        dtSettings._originalFnRecordsTotal = dtSettings.fnRecordsTotal;
                        dtSettings._originalFnRecordsDisplay = dtSettings.fnRecordsDisplay;
                    }
                    dtSettings.fnRecordsTotal = function () { return totalRecords; };
                    dtSettings.fnRecordsDisplay = function () { return totalRecords; };

                    // Update the pagination info display manually
                    let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                    if ($info.length) {
                        $info.text(`Showing ${start} to ${end} of ${totalRecords} entries`);
                    }

                    // Disable/enable last page button based on batch loading status
                    let $pagination = $(api.table().node()).parents('.table-responsive').find('.dt-paging');
                    let $lastButton = $pagination.find('button.dt-paging-button.last');
                    let $lastPageBtn = $pagination.find(
                        `button.dt-paging-button[data-dt-idx="${lastPageIndex}"]`
                    );
                    if ($lastButton.length) {
                        // Check if all batches are loaded
                        let totalBatchesNeeded = Math.ceil(totalRecords / RECORDS_PER_BATCH);
                        let allBatchesLoaded = loadedResolvedBatches >= totalBatchesNeeded || (loadedResolvedBatches * RECORDS_PER_BATCH >= totalRecords);

                        if (!allBatchesLoaded) {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .addClass('disabled')
                                    .attr('aria-disabled', 'true')
                                    .attr('tabindex', '-1')
                                    .prop('disabled', true)
                                    .css({
                                        'pointer-events': 'none',
                                        'opacity': '0.5'
                                    });
                            });
                        } else {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .removeClass('disabled')
                                    .attr('aria-disabled', 'false')
                                    .removeAttr('tabindex')
                                    .prop('disabled', false)
                                    .css({
                                        'pointer-events': '',
                                        'opacity': '1'
                                    });
                            });
                        }
                    }
                }

                // Check if we need to load next batch based on current page
                // Use debounce to prevent multiple rapid calls
                let currentPage = pageInfo.page + 1; // DataTables uses 0-based indexing
                if (!isLoadingResolvedBatch && currentPage !== lastCheckedResolvedPage) {
                    // Clear any pending timeout
                    if (allResolvedBatchLoadTimeout) {
                        clearTimeout(allResolvedBatchLoadTimeout);
                    }

                    let recordsPerPage = pageInfo.length;
                    let recordsNeeded = currentPage * recordsPerPage;
                    let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                    // Check if we need to load batches BEFORE debounce (to show loader immediately)
                    // Load if we need more batches AND we have more records than currently loaded
                    if (batchNeeded > loadedResolvedBatches && loadedResolvedBatches * RECORDS_PER_BATCH < allResolvedRequests.length && !isLoadingResolvedBatch) {
                        // Show loader IMMEDIATELY (synchronously, before any async operations)
                        let $loader = $('.custom-loader');
                        if ($loader.length) {
                            $loader.css("display", "flex");
                        }

                        // Hide "no records found" message immediately
                        let $table = $(api.table().node());
                        let $emptyRow = $table.find('tbody tr.dataTables_empty');
                        if ($emptyRow.length) {
                            $emptyRow.hide();
                        }
                    }

                    // Debounce the batch loading check
                    allResolvedBatchLoadTimeout = setTimeout(() => {
                        lastCheckedResolvedPage = currentPage;
                        let recordsPerPage = pageInfo.length;
                        let recordsNeeded = currentPage * recordsPerPage;
                        let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                        // Load all batches up to the needed one
                        // Load if we need more batches AND we have more records than currently loaded
                        if (batchNeeded > loadedResolvedBatches && loadedResolvedBatches * RECORDS_PER_BATCH < allResolvedRequests.length && !isLoadingResolvedBatch) {
                            // Save current page to restore after loading
                            let targetPage = currentPage - 1; // DataTables uses 0-based indexing

                            // Get loader reference (already shown above)
                            let $loader = $('.custom-loader');

                            // Hide "no records found" message during loading (in case it appeared)
                            let $table = $(api.table().node());
                            let $emptyRow = $table.find('tbody tr.dataTables_empty');
                            if ($emptyRow.length) {
                                $emptyRow.hide();
                            }

                            // Load batches asynchronously
                            (async () => {
                                try {
                                    while (loadedResolvedBatches < batchNeeded && !isLoadingResolvedBatch) {
                                        // Don't show loader for each individual batch, only show it once at the start
                                        if (typeof loadNextResolvedBatch === 'function') {
                                            await loadNextResolvedBatch(false);
                                        } else {
                                            console.error('loadNextResolvedBatch is not accessible!');
                                            break;
                                        }
                                    }

                                    // After loading, ensure pagination is updated
                                    let dtSettings = api.settings()[0];
                                    if (dtSettings) {
                                        dtSettings._iRecordsTotal = allResolvedRequests.length;
                                        dtSettings._iRecordsDisplay = allResolvedRequests.length;
                                    }

                                    // Restore the page the user was on (use targetPage which is the page they clicked)
                                    // Use setTimeout to ensure data is fully appended before changing page
                                    setTimeout(() => {
                                        api.page(targetPage).draw('page');
                                    }, 100);
                                } finally {
                                    // Hide loader after everything is done
                                    if ($loader.length) {
                                        $loader.css("display", "none");
                                    }
                                }
                            })();
                        }
                    }, 100); // 100ms debounce
                }
            }

            // Update header text
            $('h3.grcApproveReqCount').text(`Request ${start} - ${end} of ${totalRecords || total}`);
        },
        createdRow: function (row, data, dataIndex) {
            $('td', row).each(function (index) {
                // Add title for ellipsis columns (Vendor Name and Legal Entity)
                if (index === 4 || index === 5 || index === 6) {
                    const cellText = $(this).text().trim();
                    if (cellText) {
                        $(this).attr('title', cellText);
                    }
                }
            });
        }
    });
    grcRejectTable = new DataTable('#grcRejectTable', {
        order: [[3, 'desc']], // Sort by date column descending (newest first)
        autoWidth: false,
        pageLength: RECORDS_PER_PAGE, // Show 10 records per page
        paging: true, // Enable pagination
        pagingType: 'full_numbers', // Show full pagination controls
        dom: '<"top d-flex justify-content-between"fB>rt<"bottom d-flex justify-content-between"lip><"clear">', // Add button container (we'll hide them)
        buttons: [
            {
                extend: 'excelHtml5',
                text: 'Export', // Won't be visible, we'll trigger programmatically
                filename: function () {
                    let now = new Date();
                    let dateStr = now.toISOString().slice(0, 19).replace(/[:T]/g, '-');
                    return 'Rejected_Requests_' + dateStr;
                },
                exportOptions: {
                    modifier: {
                        search: 'applied', // Only filtered rows
                        order: 'applied',  // Respect current sort
                        page: 'all'        // Ignore pagination, export all visible rows
                    },
                    columns: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] // Only visible columns (including new columns after Requested By)
                }
            }
        ],
        columnDefs: [
            {
                targets: '_all',
                defaultContent: ''
            },
            {
                targets: 0, // Legal Entity column index
                width: "50px", // You can set it to "20%", "10rem", etc.
                className: "min-width-col"
            },
            {
                targets: 3, // Date column
                type: 'html-num-fmt',
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: [4, 5], // Legal Entity column index
                className: "ellipsis"
            },
            {
                targets: 9, // Requested By column - hide same as pending tab
                visible: false
            },
            {
                targets: 10, // Assigned Date & Time column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: 11, // Due Date column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            { targets: [15, 16, 17, 18, 19, 20, 21], visible: false } // Hide PR Number, PO Number, and other hidden columns
        ],
        drawCallback: function (settings) {
            const api = this.api();
            const pageInfo = api.page.info();
            let lastPageIndex = pageInfo.pages - 1; // zero-based
            const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1; // Zero-based, so add 1
            const end = pageInfo.end;
            const total = pageInfo.recordsDisplay; // After filtering

            // Check if filtered data exists, use that count; otherwise use original array length
            const totalRecords = (filteredRequestData && filteredRequestData.length > 0)
                ? filteredRequestData.length
                : (allRejectedRequests.length > 0
                    ? allRejectedRequests.length
                    : pageInfo.recordsTotal);

            // Update DataTables internal settings to reflect total records for pagination
            // This makes DataTables show pagination for all records, not just loaded ones
            if (allRejectedRequests.length > 0 || (filteredRequestData && filteredRequestData.length > 0)) {
                let dtSettings = api.settings()[0];
                if (dtSettings) {
                    // Force update the record counts - this is what DataTables uses for pagination
                    // Set these BEFORE DataTables recalculates
                    dtSettings._iRecordsTotal = totalRecords;
                    dtSettings._iRecordsDisplay = totalRecords;

                    // Also override the functions that DataTables uses to get record counts
                    if (!dtSettings._originalFnRecordsTotal) {
                        dtSettings._originalFnRecordsTotal = dtSettings.fnRecordsTotal;
                        dtSettings._originalFnRecordsDisplay = dtSettings.fnRecordsDisplay;
                    }
                    dtSettings.fnRecordsTotal = function () { return totalRecords; };
                    dtSettings.fnRecordsDisplay = function () { return totalRecords; };

                    // Update the pagination info display manually
                    let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                    if ($info.length) {
                        $info.text(`Showing ${start} to ${end} of ${totalRecords} entries`);
                    }

                    // Disable/enable last page button based on batch loading status
                    let $pagination = $(api.table().node()).parents('.table-responsive').find('.dt-paging');
                    let $lastButton = $pagination.find('button.dt-paging-button.last');
                    let $lastPageBtn = $pagination.find(
                        `button.dt-paging-button[data-dt-idx="${lastPageIndex}"]`
                    );
                    if ($lastButton.length) {
                        // Check if all batches are loaded
                        let totalBatchesNeeded = Math.ceil(totalRecords / RECORDS_PER_BATCH);
                        let allBatchesLoaded = loadedRejectedBatches >= totalBatchesNeeded || (loadedRejectedBatches * RECORDS_PER_BATCH >= totalRecords);

                        if (!allBatchesLoaded) {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .addClass('disabled')
                                    .attr('aria-disabled', 'true')
                                    .attr('tabindex', '-1')
                                    .prop('disabled', true)
                                    .css({
                                        'pointer-events': 'none',
                                        'opacity': '0.5'
                                    });
                            });
                        } else {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .removeClass('disabled')
                                    .attr('aria-disabled', 'false')
                                    .removeAttr('tabindex')
                                    .prop('disabled', false)
                                    .css({
                                        'pointer-events': '',
                                        'opacity': '1'
                                    });
                            });
                        }
                    }
                }

                // Check if we need to load next batch based on current page
                // Use debounce to prevent multiple rapid calls
                let currentPage = pageInfo.page + 1; // DataTables uses 0-based indexing
                if (!isLoadingRejectedBatch && currentPage !== lastCheckedRejectedPage) {
                    // Clear any pending timeout
                    if (allRejectedBatchLoadTimeout) {
                        clearTimeout(allRejectedBatchLoadTimeout);
                    }

                    let recordsPerPage = pageInfo.length;
                    let recordsNeeded = currentPage * recordsPerPage;
                    let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                    // Check if we need to load batches BEFORE debounce (to show loader immediately)
                    // Load if we need more batches AND we have more records than currently loaded
                    if (batchNeeded > loadedRejectedBatches && loadedRejectedBatches * RECORDS_PER_BATCH < allRejectedRequests.length && !isLoadingRejectedBatch) {
                        // Show loader IMMEDIATELY (synchronously, before any async operations)
                        let $loader = $('.custom-loader');
                        if ($loader.length) {
                            $loader.css("display", "flex");
                        }

                        // Hide "no records found" message immediately
                        let $table = $(api.table().node());
                        let $emptyRow = $table.find('tbody tr.dataTables_empty');
                        if ($emptyRow.length) {
                            $emptyRow.hide();
                        }
                    }

                    // Debounce the batch loading check
                    allRejectedBatchLoadTimeout = setTimeout(() => {
                        lastCheckedRejectedPage = currentPage;
                        let recordsPerPage = pageInfo.length;
                        let recordsNeeded = currentPage * recordsPerPage;
                        let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                        // Load all batches up to the needed one
                        // Load if we need more batches AND we have more records than currently loaded
                        if (batchNeeded > loadedRejectedBatches && loadedRejectedBatches * RECORDS_PER_BATCH < allRejectedRequests.length && !isLoadingRejectedBatch) {
                            // Save current page to restore after loading
                            let targetPage = currentPage - 1; // DataTables uses 0-based indexing

                            // Get loader reference (already shown above)
                            let $loader = $('.custom-loader');

                            // Hide "no records found" message during loading (in case it appeared)
                            let $table = $(api.table().node());
                            let $emptyRow = $table.find('tbody tr.dataTables_empty');
                            if ($emptyRow.length) {
                                $emptyRow.hide();
                            }

                            // Load batches asynchronously
                            (async () => {
                                try {
                                    while (loadedRejectedBatches < batchNeeded && !isLoadingRejectedBatch) {
                                        // Don't show loader for each individual batch, only show it once at the start
                                        if (typeof loadNextRejectedBatch === 'function') {
                                            await loadNextRejectedBatch(false);
                                        } else {
                                            console.error('loadNextRejectedBatch is not accessible!');
                                            break;
                                        }
                                    }

                                    // After loading, ensure pagination is updated
                                    let dtSettings = api.settings()[0];
                                    if (dtSettings) {
                                        dtSettings._iRecordsTotal = allRejectedRequests.length;
                                        dtSettings._iRecordsDisplay = allRejectedRequests.length;
                                    }

                                    // Restore the page the user was on (use targetPage which is the page they clicked)
                                    // Use setTimeout to ensure data is fully appended before changing page
                                    setTimeout(() => {
                                        api.page(targetPage).draw('page');
                                    }, 100);
                                } finally {
                                    // Hide loader after everything is done
                                    if ($loader.length) {
                                        $loader.css("display", "none");
                                    }
                                }
                            })();
                        }
                    }, 100); // 100ms debounce
                }
            }

            // Update header text
            $('h3.grcRejectReqCount').text(`Request ${start} - ${end} of ${totalRecords || total}`);
        },
        createdRow: function (row, data, dataIndex) {
            $('td', row).each(function (index) {
                // Add title for ellipsis columns (Vendor Name and Legal Entity)
                if (index === 4 || index === 5) {
                    const cellText = $(this).text().trim();
                    if (cellText) {
                        $(this).attr('title', cellText);
                    }
                }
            });
        }
    });
    returenedRequestTable = new DataTable('#returenedRequestTable', {
        order: [[3, 'desc']], // Sort by date column descending (newest first)
        autoWidth: false,
        pageLength: RECORDS_PER_PAGE, // Show 10 records per page
        paging: true, // Enable pagination
        pagingType: 'full_numbers', // Show full pagination controls
        dom: '<"top d-flex justify-content-between"fB>rt<"bottom d-flex justify-content-between"lip><"clear">', // Add button container (we'll hide them)
        buttons: [
            {
                extend: 'excelHtml5',
                text: 'Export', // Won't be visible, we'll trigger programmatically
                filename: function () {
                    let now = new Date();
                    let dateStr = now.toISOString().slice(0, 19).replace(/[:T]/g, '-');
                    return 'Returned_Requests_' + dateStr;
                },
                exportOptions: {
                    modifier: {
                        search: 'applied', // Only filtered rows
                        order: 'applied',  // Respect current sort
                        page: 'all'        // Ignore pagination, export all visible rows
                    },
                    columns: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] // Only visible columns (including new columns after Requested By)
                }
            }
        ],
        columnDefs: [
            {
                targets: '_all',
                defaultContent: ''
            },
            {
                targets: 0, // Legal Entity column index
                width: "50px", // You can set it to "20%", "10rem", etc.
                className: "min-width-col"
            },
            {
                targets: 3, // Date column
                type: 'html-num-fmt',
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: [4, 5], // Legal Entity column index
                className: "ellipsis"
            },
            {
                targets: 9, // Requested By column - hide same as pending tab
                visible: false
            },
            {
                targets: 10, // Assigned Date & Time column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: 11, // Due Date column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            { targets: [15, 16, 17, 18, 19, 20, 21], visible: false } // Hide PR Number, PO Number, and other hidden columns
        ],
        drawCallback: function (settings) {
            const api = this.api();
            const pageInfo = api.page.info();
            let lastPageIndex = pageInfo.pages - 1; // zero-based
            const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1; // Zero-based, so add 1
            const end = pageInfo.end;
            const total = pageInfo.recordsDisplay; // After filtering

            // Check if filtered data exists, use that count; otherwise use original array length
            const totalRecords = (filteredRequestData && filteredRequestData.length > 0)
                ? filteredRequestData.length
                : (allReturnedRequests.length > 0
                    ? allReturnedRequests.length
                    : pageInfo.recordsTotal);

            // Update DataTables internal settings to reflect total records for pagination
            // This makes DataTables show pagination for all records, not just loaded ones
            if (allReturnedRequests.length > 0 || (filteredRequestData && filteredRequestData.length > 0)) {
                let dtSettings = api.settings()[0];
                if (dtSettings) {
                    // Force update the record counts - this is what DataTables uses for pagination
                    // Set these BEFORE DataTables recalculates
                    dtSettings._iRecordsTotal = totalRecords;
                    dtSettings._iRecordsDisplay = totalRecords;

                    // Also override the functions that DataTables uses to get record counts
                    if (!dtSettings._originalFnRecordsTotal) {
                        dtSettings._originalFnRecordsTotal = dtSettings.fnRecordsTotal;
                        dtSettings._originalFnRecordsDisplay = dtSettings.fnRecordsDisplay;
                    }
                    dtSettings.fnRecordsTotal = function () { return totalRecords; };
                    dtSettings.fnRecordsDisplay = function () { return totalRecords; };

                    // Update the pagination info display manually
                    let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                    if ($info.length) {
                        $info.text(`Showing ${start} to ${end} of ${totalRecords} entries`);
                    }

                    // Disable/enable last page button based on batch loading status
                    let $pagination = $(api.table().node()).parents('.table-responsive').find('.dt-paging');
                    let $lastButton = $pagination.find('button.dt-paging-button.last');
                    let $lastPageBtn = $pagination.find(
                        `button.dt-paging-button[data-dt-idx="${lastPageIndex}"]`
                    );
                    if ($lastButton.length) {
                        // Check if all batches are loaded
                        let totalBatchesNeeded = Math.ceil(totalRecords / RECORDS_PER_BATCH);
                        let allBatchesLoaded = loadedReturnedBatches >= totalBatchesNeeded || (loadedReturnedBatches * RECORDS_PER_BATCH >= totalRecords);

                        if (!allBatchesLoaded) {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .addClass('disabled')
                                    .attr('aria-disabled', 'true')
                                    .attr('tabindex', '-1')
                                    .prop('disabled', true)
                                    .css({
                                        'pointer-events': 'none',
                                        'opacity': '0.5'
                                    });
                            });
                        } else {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .removeClass('disabled')
                                    .attr('aria-disabled', 'false')
                                    .removeAttr('tabindex')
                                    .prop('disabled', false)
                                    .css({
                                        'pointer-events': '',
                                        'opacity': '1'
                                    });
                            });
                        }
                    }
                }

                // Check if we need to load next batch based on current page
                // Use debounce to prevent multiple rapid calls
                let currentPage = pageInfo.page + 1; // DataTables uses 0-based indexing
                if (!isLoadingReturnedBatch && currentPage !== lastCheckedReturnedPage) {
                    // Clear any pending timeout
                    if (allReturnedBatchLoadTimeout) {
                        clearTimeout(allReturnedBatchLoadTimeout);
                    }

                    let recordsPerPage = pageInfo.length;
                    let recordsNeeded = currentPage * recordsPerPage;
                    let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                    // Check if we need to load batches BEFORE debounce (to show loader immediately)
                    // Load if we need more batches AND we have more records than currently loaded
                    if (batchNeeded > loadedReturnedBatches && loadedReturnedBatches * RECORDS_PER_BATCH < allReturnedRequests.length && !isLoadingReturnedBatch) {
                        // Show loader IMMEDIATELY (synchronously, before any async operations)
                        let $loader = $('.custom-loader');
                        if ($loader.length) {
                            $loader.css("display", "flex");
                        }

                        // Hide "no records found" message immediately
                        let $table = $(api.table().node());
                        let $emptyRow = $table.find('tbody tr.dataTables_empty');
                        if ($emptyRow.length) {
                            $emptyRow.hide();
                        }
                    }

                    // Debounce the batch loading check
                    allReturnedBatchLoadTimeout = setTimeout(() => {
                        lastCheckedReturnedPage = currentPage;
                        let recordsPerPage = pageInfo.length;
                        let recordsNeeded = currentPage * recordsPerPage;
                        let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                        // Load all batches up to the needed one
                        // Load if we need more batches AND we have more records than currently loaded
                        if (batchNeeded > loadedReturnedBatches && loadedReturnedBatches * RECORDS_PER_BATCH < allReturnedRequests.length && !isLoadingReturnedBatch) {
                            // Save current page to restore after loading
                            let targetPage = currentPage - 1; // DataTables uses 0-based indexing

                            // Get loader reference (already shown above)
                            let $loader = $('.custom-loader');

                            // Hide "no records found" message during loading (in case it appeared)
                            let $table = $(api.table().node());
                            let $emptyRow = $table.find('tbody tr.dataTables_empty');
                            if ($emptyRow.length) {
                                $emptyRow.hide();
                            }

                            // Load batches asynchronously
                            (async () => {
                                try {
                                    while (loadedReturnedBatches < batchNeeded && !isLoadingReturnedBatch) {
                                        // Don't show loader for each individual batch, only show it once at the start
                                        if (typeof loadNextReturnedBatch === 'function') {
                                            await loadNextReturnedBatch(false);
                                        } else {
                                            console.error('loadNextReturnedBatch is not accessible!');
                                            break;
                                        }
                                    }

                                    // After loading, ensure pagination is updated
                                    let dtSettings = api.settings()[0];
                                    if (dtSettings) {
                                        dtSettings._iRecordsTotal = allReturnedRequests.length;
                                        dtSettings._iRecordsDisplay = allReturnedRequests.length;
                                    }

                                    // Restore the page the user was on (use targetPage which is the page they clicked)
                                    // Use setTimeout to ensure data is fully appended before changing page
                                    setTimeout(() => {
                                        api.page(targetPage).draw('page');
                                    }, 100);
                                } finally {
                                    // Hide loader after everything is done
                                    if ($loader.length) {
                                        $loader.css("display", "none");
                                    }
                                }
                            })();
                        }
                    }, 100); // 100ms debounce
                }
            }

            // Update header text
            $('h3.returnedReqCount').text(`Request ${start} - ${end} of ${totalRecords || total}`);
        },
        createdRow: function (row, data, dataIndex) {
            $('td', row).each(function (index) {
                // Add title for ellipsis columns (Vendor Name and Legal Entity)
                if (index === 4 || index === 5) {
                    const cellText = $(this).text().trim();
                    if (cellText) {
                        $(this).attr('title', cellText);
                    }
                }
            });
        }
    });
    reassignedTable = new DataTable('#reassignedTable', {
        order: [[3, 'desc']], // Sort by date column descending (newest first)
        autoWidth: false,
        dom: '<"top d-flex justify-content-between"fB>rt<"bottom d-flex justify-content-between"lip><"clear">', // Add button container (we'll hide them)
        buttons: [
            {
                extend: 'excelHtml5',
                text: 'Export', // Won't be visible, we'll trigger programmatically
                filename: function () {
                    let now = new Date();
                    let dateStr = now.toISOString().slice(0, 19).replace(/[:T]/g, '-');
                    return 'Re-assigned_Requests_' + dateStr;
                },
                exportOptions: {
                    modifier: {
                        search: 'applied', // Only filtered rows
                        order: 'applied',  // Respect current sort
                        page: 'all'        // Ignore pagination, export all visible rows
                    },
                    columns: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] // Only visible columns (including new columns after Requested By)
                }
            }
        ],
        columnDefs: [
            {
                targets: '_all',
                defaultContent: ''
            },
            {
                targets: 0, // Legal Entity column index
                width: "50px", // You can set it to "20%", "10rem", etc.
                className: "min-width-col"
            },
            {
                targets: 3, // Date column
                type: 'html-num-fmt',
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: [4, 5], // Legal Entity column index
                className: "ellipsis"
            },
            {
                targets: 9, // Requested By column - hide same as pending tab
                visible: false
            },
            {
                targets: 10, // Assigned Date & Time column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: 11, // Due Date column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            { targets: [15, 16, 17, 18, 19, 20, 21], visible: false } // Hide PR Number, PO Number, and other hidden columns
        ],
        drawCallback: function (settings) {
            const pageInfo = this.api().page.info();
            const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1; // Zero-based, so add 1
            const end = pageInfo.end;
            const total = pageInfo.recordsDisplay; // After filtering

            // Update header text
            $('h3.reassignedReqCount').text(`Request ${start} - ${end} of ${total}`);
        },
        createdRow: function (row, data, dataIndex) {
            $('td', row).each(function (index) {
                // Add title for ellipsis columns (Vendor Name and Legal Entity)
                if (index === 4 || index === 5) {
                    const cellText = $(this).text().trim();
                    if (cellText) {
                        $(this).attr('title', cellText);
                    }
                }
            });
        }
    });
    p1Table = new DataTable('#p1Table', {
        order: [[3, 'desc']], // Sort by date column descending (newest first)
        autoWidth: false,
        pageLength: RECORDS_PER_PAGE, // Show 10 records per page
        paging: true, // Enable pagination
        pagingType: 'full_numbers', // Show full pagination controls
        dom: '<"top d-flex justify-content-between"fB>rt<"bottom d-flex justify-content-between"lip><"clear">', // Add button container (we'll hide them)
        buttons: [
            {
                extend: 'excelHtml5',
                text: 'Export', // Won't be visible, we'll trigger programmatically
                filename: function () {
                    let now = new Date();
                    let dateStr = now.toISOString().slice(0, 19).replace(/[:T]/g, '-');
                    return 'P1_Requests_' + dateStr;
                },
                exportOptions: {
                    modifier: {
                        search: 'applied', // Only filtered rows
                        order: 'applied',  // Respect current sort
                        page: 'all'        // Ignore pagination, export all visible rows
                    },
                    columns: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] // Only visible columns (including new columns after Requested By)
                }
            }
        ],
        columnDefs: [
            {
                targets: '_all',
                defaultContent: ''
            },
            {
                targets: 0, // Legal Entity column index
                width: "50px", // You can set it to "20%", "10rem", etc.
                className: "min-width-col"
            },
            {
                targets: 3, // Date column
                type: 'html-num-fmt',
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: [4, 5], // Legal Entity column index
                className: "ellipsis"
            },
            {
                targets: 9, // Requested By column - hide same as pending tab
                visible: false
            },
            {
                targets: 10, // Assigned Date & Time column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            {
                targets: 11, // Due Date column - hide same as pending tab
                type: 'html-num-fmt',
                visible: false,
                render: function (data, type, row) {
                    if (type === 'sort' || type === 'type') {
                        return parseDateForSorting(data);
                    }
                    return data;
                }
            },
            { targets: [15, 16, 17, 18, 19, 20, 21], visible: false } // Hide PR Number, PO Number, and other hidden columns
        ],
        drawCallback: function (settings) {
            const api = this.api();
            const pageInfo = api.page.info();
            let lastPageIndex = pageInfo.pages - 1; // zero-based
            const start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1; // Zero-based, so add 1
            const end = pageInfo.end;
            const total = pageInfo.recordsDisplay; // After filtering

            // Check if filtered data exists, use that count; otherwise use original array length
            const totalRecords = (filteredRequestData && filteredRequestData.length > 0)
                ? filteredRequestData.length
                : (allP1Requests.length > 0
                    ? allP1Requests.length
                    : pageInfo.recordsTotal);

            // Update DataTables internal settings to reflect total records for pagination
            // This makes DataTables show pagination for all records, not just loaded ones
            if (allP1Requests.length > 0 || (filteredRequestData && filteredRequestData.length > 0)) {
                let dtSettings = api.settings()[0];
                if (dtSettings) {
                    // Force update the record counts - this is what DataTables uses for pagination
                    // Set these BEFORE DataTables recalculates
                    dtSettings._iRecordsTotal = totalRecords;
                    dtSettings._iRecordsDisplay = totalRecords;

                    // Also override the functions that DataTables uses to get record counts
                    if (!dtSettings._originalFnRecordsTotal) {
                        dtSettings._originalFnRecordsTotal = dtSettings.fnRecordsTotal;
                        dtSettings._originalFnRecordsDisplay = dtSettings.fnRecordsDisplay;
                    }
                    dtSettings.fnRecordsTotal = function () { return totalRecords; };
                    dtSettings.fnRecordsDisplay = function () { return totalRecords; };

                    // Update the pagination info display manually
                    let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                    if ($info.length) {
                        $info.text(`Showing ${start} to ${end} of ${totalRecords} entries`);
                    }

                    // Disable/enable last page button based on batch loading status
                    let $pagination = $(api.table().node()).parents('.table-responsive').find('.dt-paging');
                    let $lastButton = $pagination.find('button.dt-paging-button.last');
                    let $lastPageBtn = $pagination.find(
                        `button.dt-paging-button[data-dt-idx="${lastPageIndex}"]`
                    );
                    if ($lastButton.length) {
                        // Check if all batches are loaded
                        let totalBatchesNeeded = Math.ceil(totalRecords / RECORDS_PER_BATCH);
                        let allBatchesLoaded = loadedP1Batches >= totalBatchesNeeded || (loadedP1Batches * RECORDS_PER_BATCH >= totalRecords);

                        if (!allBatchesLoaded) {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .addClass('disabled')
                                    .attr('aria-disabled', 'true')
                                    .attr('tabindex', '-1')
                                    .prop('disabled', true)
                                    .css({
                                        'pointer-events': 'none',
                                        'opacity': '0.5'
                                    });
                            });
                        } else {
                            [$lastButton, $lastPageBtn].forEach($btn => {
                                $btn
                                    .removeClass('disabled')
                                    .attr('aria-disabled', 'false')
                                    .removeAttr('tabindex')
                                    .prop('disabled', false)
                                    .css({
                                        'pointer-events': '',
                                        'opacity': '1'
                                    });
                            });
                        }
                    }
                }

                // Check if we need to load next batch based on current page
                // Use debounce to prevent multiple rapid calls
                let currentPage = pageInfo.page + 1; // DataTables uses 0-based indexing
                if (!isLoadingP1Batch && currentPage !== lastCheckedP1Page) {
                    // Clear any pending timeout
                    if (allP1BatchLoadTimeout) {
                        clearTimeout(allP1BatchLoadTimeout);
                    }

                    let recordsPerPage = pageInfo.length;
                    let recordsNeeded = currentPage * recordsPerPage;
                    let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                    // Check if we need to load batches BEFORE debounce (to show loader immediately)
                    if (batchNeeded > loadedP1Batches && batchNeeded * RECORDS_PER_BATCH <= allP1Requests.length && !isLoadingP1Batch) {
                        // Show loader IMMEDIATELY (synchronously, before any async operations)
                        let $loader = $('.custom-loader');
                        if ($loader.length) {
                            $loader.css("display", "flex");
                        }

                        // Hide "no records found" message immediately
                        let $table = $(api.table().node());
                        let $emptyRow = $table.find('tbody tr.dataTables_empty');
                        if ($emptyRow.length) {
                            $emptyRow.hide();
                        }
                    }

                    // Debounce the batch loading check
                    allP1BatchLoadTimeout = setTimeout(() => {
                        lastCheckedP1Page = currentPage;
                        let recordsPerPage = pageInfo.length;
                        let recordsNeeded = currentPage * recordsPerPage;
                        let batchNeeded = Math.ceil(recordsNeeded / RECORDS_PER_BATCH);

                        // Load all batches up to the needed one
                        if (batchNeeded > loadedP1Batches && batchNeeded * RECORDS_PER_BATCH <= allP1Requests.length && !isLoadingP1Batch) {
                            // Save current page to restore after loading
                            let targetPage = currentPage - 1; // DataTables uses 0-based indexing

                            // Get loader reference (already shown above)
                            let $loader = $('.custom-loader');

                            // Hide "no records found" message during loading (in case it appeared)
                            let $table = $(api.table().node());
                            let $emptyRow = $table.find('tbody tr.dataTables_empty');
                            if ($emptyRow.length) {
                                $emptyRow.hide();
                            }

                            // Load batches asynchronously
                            (async () => {
                                try {
                                    while (loadedP1Batches < batchNeeded && !isLoadingP1Batch) {
                                        // Don't show loader for each individual batch, only show it once at the start
                                        if (typeof loadNextP1Batch === 'function') {
                                            await loadNextP1Batch(false);
                                        } else {
                                            console.error('loadNextP1Batch is not accessible!');
                                            break;
                                        }
                                    }

                                    // After loading, ensure pagination is updated
                                    let dtSettings = api.settings()[0];
                                    if (dtSettings) {
                                        dtSettings._iRecordsTotal = allP1Requests.length;
                                        dtSettings._iRecordsDisplay = allP1Requests.length;
                                    }

                                    // Restore the page the user was on (use targetPage which is the page they clicked)
                                    // Use setTimeout to ensure data is fully appended before changing page
                                    setTimeout(() => {
                                        api.page(targetPage).draw('page');
                                    }, 100);
                                } finally {
                                    // Hide loader after everything is done
                                    if ($loader.length) {
                                        $loader.css("display", "none");
                                    }
                                }
                            })();
                        }
                    }, 100); // 100ms debounce
                }
            }

            // Update header text
            $('h3.p1ReqCount').text(`Request ${start} - ${end} of ${totalRecords || total}`);
        },
        createdRow: function (row, data, dataIndex) {
            $('td', row).each(function (index) {
                // Add title for ellipsis columns (Vendor Name and Legal Entity)
                if (index === 4 || index === 5) {
                    const cellText = $(this).text().trim();
                    if (cellText) {
                        $(this).attr('title', cellText);
                    }
                }
            });
        }
    });

    // Hide the built-in export button
    allReqTable.buttons().container().hide();
    pendingReqTable.buttons().container().hide();
    grcApproveTable.buttons().container().hide();
    grcRejectTable.buttons().container().hide();
    returenedRequestTable.buttons().container().hide();
    reassignedTable.buttons().container().hide();
    p1Table.buttons().container().hide();

    // Trigger download when your header link is clicked
    $('.card_table_list_filter a:contains("Download")').on('click', function (e) {
        e.preventDefault();
        if ($('#pendingReqTable').is(":visible")) {
            pendingReqTable.button(0).trigger(); // Trigger Excel export
        } else if ($('#grcApproveTable').is(":visible")) {
            grcApproveTable.button(0).trigger(); // Trigger Excel export
        }
        else if ($('#grcRejectTable').is(":visible")) {
            grcRejectTable.button(0).trigger(); // Trigger Excel export
        }
        else if ($('#returenedRequestTable').is(":visible")) {
            returenedRequestTable.button(0).trigger(); // Trigger Excel export
        }
        else if ($('#reassignedTable').is(":visible")) {
            reassignedTable.button(0).trigger();
        }
        else if ($('#p1Table').is(":visible")) {
            p1Table.button(0).trigger();
        }
        else {
            allReqTable.button(0).trigger(); // Trigger Excel export
        }

    });

    (function (webapi, $) {
        function safeAjax(ajaxOptions) {
            var deferredAjax = $.Deferred();

            shell.getTokenDeferred().done(function (token) {
                // add headers for AJAX
                if (!ajaxOptions.headers) {
                    $.extend(ajaxOptions, {
                        headers: {
                            "__RequestVerificationToken": token
                        }
                    });
                } else {
                    ajaxOptions.headers["__RequestVerificationToken"] = token;
                }
                $.ajax(ajaxOptions)
                    .done(function (data, textStatus, jqXHR) {
                        validateLoginSession(data, textStatus, jqXHR, deferredAjax.resolve);
                    }).fail(deferredAjax.reject); //AJAX
            }).fail(function () {
                deferredAjax.rejectWith(this, arguments); // on token failure pass the token AJAX and args
            });

            return deferredAjax.promise();
        }
        webapi.safeAjax = safeAjax;
    })(window.webapi = window.webapi || {}, jQuery);

    await onLoad('Pending');
    if (window.location.pathname === "/my-request/") {
        setInterval(async function () {
            await onLoad('Pending');
        }, 300000); // 300,000 ms = 300 seconds = 5 minutes
    }
});

async function getApprovalRequests() {
    return new Promise((resolve, reject) => {
        webapi.safeAjax({
            type: "GET",
            url: "/_api/al_approvalrequests?$select=al_requeststatus,al_finalrequeststatus,_al_request_value,_al_team_value&$orderby=modifiedon desc",
            contentType: "application/json",
            headers: { "Prefer": "odata.include-annotations=*" },
            success: function (data) {
                approvalRequest = data.value;
                resolve(approvalRequest);
            },
            error: function (xhr) {
                reject(xhr);
                console.error("Failed checking request:", xhr);
            }
        });
    });
}

$(document).on("click", "#myTab .nav-link", function () {
    var elem = $(this).text();
    onLoad(elem);
});

async function onLoad(elem) {
    $('.custom-loader').css("display", "flex");
    loggedInUserId = $("#userIdHidden").html();
    if (loggedInUserId) {
        try {
            //Get User Team Id
            userWithTeam = await GetUserWithTeam(loggedInUserId);
            console.log(userWithTeam);
            if (!userWithTeam) {
                showErrorModal("Error fetching team details.");
                $('.custom-loader').css("display", "none");
                return;
            }
            let teamType = userWithTeam.al_teamtype;
            currentUserTeamType = teamType;
            let myRequestList = [];

            // Fetch approval requests - don't let this fail stop everything
            try {
                //approvalRequest = await getApprovalRequests();
            } catch (error) {
                console.error("Error fetching approval requests:", error);
                approvalRequest = [];
            }

            // Fetch P1 requests - don't let this fail stop everything
            try {
                if (elem && (elem.includes('P1') || elem === 'p1')) {
                    let p1Requests = await GetP1Requests(loggedInUserId);
                    if (p1Requests && p1Requests.length > 0) {
                        await BindRequest(userWithTeam, p1Requests, p1Table, teamType);
                    }
                }
            } catch (error) {
                console.error("Error loading P1 requests:", error);
            }

            try {
                switch (parseInt(teamType)) {
                    case 1://1 BU
                        $('#reassigned-tab').hide();

                        // Get the appropriate table based on clicked tab
                        let targetTable = getTableByTabName(elem);
                        let allRecords = [];
                        let loadedBatchesRef = { value: 0 };
                        let tableInstance = null;
                        // Get ALL requests first for KPI counts (regardless of which tab is clicked)
                        allRequestsForKPICounts = await GetBuTeamRequests(loggedInUserId, "All");
                        // Directly set total count in KPI card immediately after API call
                        if (allRequestsForKPICounts && allRequestsForKPICounts.length > 0) {
                            $(".totalReq").text(allRequestsForKPICounts.length);
                        }
                        // Determine which array to use based on tab name
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            // All tab - store in allAllRequests
                            allAllRequests = await GetBuTeamRequests(loggedInUserId, elem);
                            allRecords = allAllRequests;
                            loadedBatchesRef = { value: loadedAllBatches };
                            tableInstance = allReqTable;
                            loadedAllBatches = 0;
                            allRequestDataStorage = allRecords;
                        } else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            // Returned tab
                            allReturnedRequests = await GetBuTeamRequests(loggedInUserId, elem);
                            allRecords = allReturnedRequests;
                            loadedBatchesRef = { value: loadedReturnedBatches };
                            tableInstance = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            allRequestDataStorage = allRecords;
                        } else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'approved requests' || elem === 'Resolved')) {
                            // Resolved/Approved tab
                            allResolvedRequests = await GetBuTeamRequests(loggedInUserId, elem);
                            allRecords = allResolvedRequests;
                            loadedBatchesRef = { value: loadedResolvedBatches };
                            tableInstance = grcApproveTable;
                            loadedResolvedBatches = 0;
                            allRequestDataStorage = allRecords;
                        } else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            // Rejected tab
                            allRejectedRequests = await GetBuTeamRequests(loggedInUserId, elem);
                            allRecords = allRejectedRequests;
                            loadedBatchesRef = { value: loadedRejectedBatches };
                            tableInstance = grcRejectTable;
                            loadedRejectedBatches = 0;
                            allRequestDataStorage = allRecords;
                        } else {
                            // Default to Pending tab
                            allPendingRequests = await GetBuTeamRequests(loggedInUserId, elem);
                            allRecords = allPendingRequests;
                            loadedBatchesRef = { value: loadedPendingBatches };
                            tableInstance = pendingReqTable;
                            loadedPendingBatches = 0;
                            $(".pendingReq").text(allPendingRequests?.length || 0);
                            // Store all data for filtering functionality
                            allRequestDataStorage = allPendingRequests;
                            filteredRequestData = null; // Reset filtered data
                        }
                        // Update KPI cards with total counts from ALL requests (not just current tab)
                        if (allRequestsForKPICounts && allRequestsForKPICounts.length > 0) {
                            getRequestCount(allRequestsForKPICounts);

                        }
                        // Load first batch and update pagination
                        if (allRecords.length > 0 && tableInstance) {
                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecords,
                                loadedBatches: loadedBatchesRef,
                                tableInstance: tableInstance,
                                bindRequestTeamType: -1
                            });
                            loadedBatchesRef.value = 1;

                            // Update the appropriate batch counter
                            if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                                loadedAllBatches = 1;
                            } else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                                loadedReturnedBatches = 1;
                            } else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'approved requests' || elem === 'Resolved')) {
                                loadedResolvedBatches = 1;
                            } else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                                loadedRejectedBatches = 1;
                            } else {
                                loadedPendingBatches = 1;
                            }

                            // Update DataTables settings to show pagination for all records
                            setTimeout(() => {
                                if (tableInstance && allRecords.length > 0) {
                                    try {
                                        if (typeof tableInstance.api === 'function') {
                                            let api = tableInstance.api();
                                            if (api && api.settings && api.settings().length > 0) {
                                                let dtSettings = api.settings()[0];
                                                if (dtSettings) {
                                                    dtSettings._iRecordsTotal = allRecords.length;
                                                    dtSettings._iRecordsDisplay = allRecords.length;
                                                    dtSettings.fnRecordsTotal = function () { return allRecords.length; };
                                                    dtSettings.fnRecordsDisplay = function () { return allRecords.length; };
                                                    api.draw(false);
                                                }
                                            }
                                        }
                                    } catch (error) {
                                        console.warn('Error updating pagination settings:', error);
                                    }
                                }
                            }, 100);
                        }

                        // //Returned Requests 
                        // let returnedRequestList = myRequestList.filter(x => x.al_requeststatus == 5);
                        // await BindRequest(userWithTeam, returnedRequestList, returenedRequestTable, teamType);

                        // //Approved Requests //Resolved & Resolved UVP
                        // let approvedRequestList = myRequestList.filter(x => x.al_requeststatus == 6 || x.al_requeststatus == 40);
                        // await BindRequest(userWithTeam, approvedRequestList, grcApproveTable, teamType);

                        // //Rejected Requests
                        // let rejectedRequestList = myRequestList.filter(x => x.al_requeststatus == 4);
                        // await BindRequest(userWithTeam, rejectedRequestList, grcRejectTable, teamType);

                        // $(".kpiCards").removeClass("d-none");
                        // getRequestCount(myRequestList);
                        // await getNotificationItems()
                        $(".kpiCards").show();
                        $(".kpiCards").css("display", "");
                        $(".kpiCards").removeClass("d-none");
                        break;
                    case 2:// 2 Line Manager
                        $('#reassigned-tab').hide();
                        $("#registerNewReqBtn").hide();

                        // Get ALL requests for KPI counts (regardless of which tab is clicked)
                        allRequestsForKPICounts = await GetManagerRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;
                        // Directly set total count in KPI card immediately after API call

                        if (allRequestsForKPICounts && allRequestsForKPICounts.length > 0) {

                            $(".totalReq").text(allRequestsForKPICounts.length);

                        }

                        // Get the appropriate table based on clicked tab
                        let targetTableLM = getTableByTabName(elem);
                        let allRecordsLM = [];
                        let loadedBatchesRefLM = { value: 0 };
                        let tableInstanceLM = null;

                        // Determine which array to use based on tab name
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            // All tab - store in allAllRequests
                            allAllRequests = [...myRequestList];
                            allRecordsLM = allAllRequests;
                            loadedBatchesRefLM = { value: loadedAllBatches };
                            tableInstanceLM = allReqTable;
                            loadedAllBatches = 0;
                            allRequestDataStorage = allRecordsLM;
                        } else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {
                            // Pending tab
                            let lineManagerPendingRequests = myRequestList.filter(x => ((x.al_requeststatus == 2 || x.al_requeststatus == 41) && x._al_primarycontact_value == loggedInUserId) || (x._al_primarycontact_value == loggedInUserId && x.al_requeststatus == 5));
                            allPendingRequests = [...lineManagerPendingRequests];
                            allRecordsLM = allPendingRequests;
                            loadedBatchesRefLM = { value: loadedPendingBatches };
                            tableInstanceLM = pendingReqTable;
                            loadedPendingBatches = 0;
                            allRequestDataStorage = allRecordsLM;
                            $(".pendingReq").text(allPendingRequests?.length || 0);
                        } else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            // Returned tab
                            let returnedRequestListLM = myRequestList.filter(x => x.al_requeststatus == 5);
                            allReturnedRequests = [...returnedRequestListLM];
                            allRecordsLM = allReturnedRequests;
                            loadedBatchesRefLM = { value: loadedReturnedBatches };
                            tableInstanceLM = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            allRequestDataStorage = allRecordsLM;
                        } else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'approved requests' || elem === 'Resolved')) {
                            // Resolved/Approved tab
                            let approvedRequestListLM = myRequestList.filter(x => x.al_requeststatus == 6 || x.al_requeststatus == 40);
                            allResolvedRequests = [...approvedRequestListLM];
                            allRecordsLM = allResolvedRequests;
                            loadedBatchesRefLM = { value: loadedResolvedBatches };
                            tableInstanceLM = grcApproveTable;
                            loadedResolvedBatches = 0;
                            allRequestDataStorage = allRecordsLM;
                        } else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            // Rejected tab
                            let rejectedRequestListLM = myRequestList.filter(x => x.al_requeststatus == 4);
                            allRejectedRequests = [...rejectedRequestListLM];
                            allRecordsLM = allRejectedRequests;
                            loadedBatchesRefLM = { value: loadedRejectedBatches };
                            tableInstanceLM = grcRejectTable;
                            loadedRejectedBatches = 0;
                            allRequestDataStorage = allRecordsLM;
                        }

                        // Load first batch and update pagination
                        if (allRecordsLM.length > 0 && tableInstanceLM) {
                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsLM,
                                loadedBatches: loadedBatchesRefLM,
                                tableInstance: tableInstanceLM,
                                bindRequestTeamType: teamType
                            });
                            loadedBatchesRefLM.value = 1;

                            // Update the appropriate batch counter
                            if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                                loadedAllBatches = 1;
                            } else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {
                                loadedPendingBatches = 1;
                            } else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                                loadedReturnedBatches = 1;
                            } else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'approved requests' || elem === 'Resolved')) {
                                loadedResolvedBatches = 1;
                            } else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                                loadedRejectedBatches = 1;
                            }

                            // Update DataTables settings to show pagination for all records
                            setTimeout(() => {
                                if (tableInstanceLM && allRecordsLM.length > 0) {
                                    try {
                                        if (typeof tableInstanceLM.api === 'function') {
                                            let api = tableInstanceLM.api();
                                            if (api && api.settings && api.settings().length > 0) {
                                                let dtSettings = api.settings()[0];
                                                if (dtSettings) {
                                                    dtSettings._iRecordsTotal = allRecordsLM.length;
                                                    dtSettings._iRecordsDisplay = allRecordsLM.length;
                                                    dtSettings.fnRecordsTotal = function () { return allRecordsLM.length; };
                                                    dtSettings.fnRecordsDisplay = function () { return allRecordsLM.length; };
                                                    api.draw(false);
                                                }
                                            }
                                        }
                                    } catch (error) {
                                        console.warn('Error updating pagination settings:', error);
                                    }
                                }
                            }, 100);
                        }

                        $(".kpiCards").removeClass("d-none");
                        // Use allRequestsForKPICounts for KPI counts (all requests, not filtered by tab)

                        if (allRequestsForKPICounts && allRequestsForKPICounts.length > 0) {

                            getRequestCount(allRequestsForKPICounts);

                        }
                        break;
                    case 3: // 3 PSSC
                        $("#registerNewReqBtn").hide();
                        $("#assignBtn").removeClass("d-none");
                        $('#reassigned-tab').show();

                        // Get ALL requests once (for KPIs + all tabs)
                        allRequestsForKPICounts = await GetAllRequestForPSSC(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        // KPI – total
                        $(".totalReq").text(myRequestList?.length || 0);

                        // Get table based on clicked tab
                        let targetTablePSSC = getTableByTabName(elem);
                        let allRecordsPSSC = [];
                        let loadedBatchesRefPSSC = { value: 0 };
                        let tableInstancePSSC = null;

                        // -------- TAB LOGIC (same pattern as Line Manager) --------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsPSSC = allAllRequests;
                            tableInstancePSSC = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefPSSC = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsPSSC;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {

                            let psscPendingRequests = myRequestList.filter(x => {
                                const latestPsscApproval = (x.al_approvalrequest_Request_al_request || [])
                                    .filter(a => a.al_teamtype === 3)
                                    .sort((a, b) => new Date(b.createdon) - new Date(a.createdon))[0];

                                return (
                                    (x.al_requeststatus === 3 && x._al_primarycontact_value === loggedInUserId) ||
                                    (x.al_requeststatus === 50 && x._al_primarycontact_value === loggedInUserId) ||
                                    x.al_requeststatus === 41 ||
                                    (x.al_requeststatus === 5 && x._al_primarycontact_value === loggedInUserId) ||
                                    (x.al_requeststatus === 49 && latestPsscApproval?.al_requeststatus === 1)
                                );
                            });

                            allPendingRequests = [...psscPendingRequests];
                            allRecordsPSSC = allPendingRequests;
                            tableInstancePSSC = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefPSSC = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsPSSC;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            allReturnedRequests = myRequestList.filter(x => x.al_requeststatus === 5);
                            allRecordsPSSC = allReturnedRequests;
                            tableInstancePSSC = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefPSSC = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsPSSC;
                        }

                        // RESOLVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Resolved')) {
                            allResolvedRequests = myRequestList.filter(
                                x => x.al_requeststatus === 6 || x.al_requeststatus === 40
                            );
                            allRecordsPSSC = allResolvedRequests;
                            tableInstancePSSC = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefPSSC = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsPSSC;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            allRejectedRequests = myRequestList.filter(x => x.al_requeststatus === 4);
                            allRecordsPSSC = allRejectedRequests;
                            tableInstancePSSC = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefPSSC = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsPSSC;
                        }

                        // -------- LOAD FIRST BATCH (same as Line Manager) --------

                        if (allRecordsPSSC.length > 0 && tableInstancePSSC) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsPSSC,
                                loadedBatches: loadedBatchesRefPSSC,
                                tableInstance: tableInstancePSSC,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefPSSC.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('Resolved') || elem.includes('approved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination
                            setTimeout(() => {
                                try {
                                    let api = tableInstancePSSC.api();
                                    let settings = api.settings()[0];
                                    settings._iRecordsTotal = allRecordsPSSC.length;
                                    settings._iRecordsDisplay = allRecordsPSSC.length;
                                    settings.fnRecordsTotal = () => allRecordsPSSC.length;
                                    settings.fnRecordsDisplay = () => allRecordsPSSC.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // KPIs
                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);

                        break;

                    case 4: // 4 GRC
                        $("#reassigned-tab").hide();
                        $("#registerNewReqBtn").hide();

                        // Get ALL requests (used for KPIs + all tabs)
                        allRequestsForKPICounts = await GetComplianceRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        // KPI – total
                        $(".totalReq").text(myRequestList?.length || 0);

                        // Table handling (same as Line Manager)
                        let targetTableGRC = getTableByTabName(elem);
                        let allRecordsGRC = [];
                        let loadedBatchesRefGRC = { value: 0 };
                        let tableInstanceGRC = null;

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsGRC = allAllRequests;
                            tableInstanceGRC = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefGRC = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsGRC;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {

                            let grcPendingRequests = myRequestList.filter(x =>
                                (x.al_requeststatus === 7 || x.al_requeststatus === 41) ||
                                (x._al_primarycontact_value === loggedInUserId && x.al_requeststatus === 5)
                            );

                            allPendingRequests = [...grcPendingRequests];
                            allRecordsGRC = allPendingRequests;
                            tableInstanceGRC = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefGRC = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsGRC;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            allReturnedRequests = myRequestList.filter(x => x.al_requeststatus === 5);
                            allRecordsGRC = allReturnedRequests;
                            tableInstanceGRC = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefGRC = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsGRC;
                        }

                        // APPROVED / RESOLVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Approved')) {
                            allResolvedRequests = [...grcApproveRequestList];
                            allRecordsGRC = allResolvedRequests;
                            tableInstanceGRC = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefGRC = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsGRC;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            allRejectedRequests = [...grcRejectRequestList];
                            allRecordsGRC = allRejectedRequests;
                            tableInstanceGRC = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefGRC = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsGRC;
                        }

                        // ---------------- LOAD FIRST BATCH ----------------

                        if (allRecordsGRC.length > 0 && tableInstanceGRC) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsGRC,
                                loadedBatches: loadedBatchesRefGRC,
                                tableInstance: tableInstanceGRC,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefGRC.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('approved') || elem.includes('Resolved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination (same as LM)
                            setTimeout(() => {
                                try {
                                    let api = tableInstanceGRC.api();
                                    let settings = api.settings()[0];
                                    settings._iRecordsTotal = allRecordsGRC.length;
                                    settings._iRecordsDisplay = allRecordsGRC.length;
                                    settings.fnRecordsTotal = () => allRecordsGRC.length;
                                    settings.fnRecordsDisplay = () => allRecordsGRC.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // ---------------- KPI / CARDS ----------------

                        $("#grc-approve-tab").html(`Approved : <b>${grcApproveRequestList.length}</b>`);
                        $("#grc-reject-tab").html(`Rejected : <b>${grcRejectRequestList.length}</b>`);

                        $(".compliance-count").removeClass("d-none");
                        $(".grcCards").removeClass("d-none");

                        getGRCCardsCount(allRequestsForKPICounts);

                        break;

                    case 5: // 5 SVP
                        $("#reassigned-tab").hide();

                        // Get ALL requests once (for KPIs + all tabs)
                        allRequestsForKPICounts = await GetManagerRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        // KPI – total
                        $(".totalReq").text(myRequestList?.length || 0);

                        // Table handling (same pattern as LM)
                        let targetTableSVP = getTableByTabName(elem);
                        let allRecordsSVP = [];
                        let loadedBatchesRefSVP = { value: 0 };
                        let tableInstanceSVP = null;

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsSVP = allAllRequests;
                            tableInstanceSVP = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefSVP = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsSVP;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {

                            let svpPendingRequests = myRequestList.filter(x =>
                                (x.al_requeststatus === 15 || x.al_requeststatus === 41) ||
                                (x._al_primarycontact_value === loggedInUserId && x.al_requeststatus === 5)
                            );

                            allPendingRequests = [...svpPendingRequests];
                            allRecordsSVP = allPendingRequests;
                            tableInstanceSVP = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefSVP = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsSVP;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            allReturnedRequests = myRequestList.filter(x => x.al_requeststatus === 5);
                            allRecordsSVP = allReturnedRequests;
                            tableInstanceSVP = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefSVP = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsSVP;
                        }

                        // RESOLVED / APPROVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Resolved')) {
                            allResolvedRequests = myRequestList.filter(
                                x => x.al_requeststatus === 6 || x.al_requeststatus === 40
                            );
                            allRecordsSVP = allResolvedRequests;
                            tableInstanceSVP = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefSVP = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsSVP;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            allRejectedRequests = myRequestList.filter(x => x.al_requeststatus === 4);
                            allRecordsSVP = allRejectedRequests;
                            tableInstanceSVP = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefSVP = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsSVP;
                        }

                        // ---------------- LOAD FIRST BATCH ----------------

                        if (allRecordsSVP.length > 0 && tableInstanceSVP) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsSVP,
                                loadedBatches: loadedBatchesRefSVP,
                                tableInstance: tableInstanceSVP,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefSVP.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('approved') || elem.includes('Resolved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination (same as LM)
                            setTimeout(() => {
                                try {
                                    let api = tableInstanceSVP.api();
                                    let settings = api.settings()[0];
                                    settings._iRecordsTotal = allRecordsSVP.length;
                                    settings._iRecordsDisplay = allRecordsSVP.length;
                                    settings.fnRecordsTotal = () => allRecordsSVP.length;
                                    settings.fnRecordsDisplay = () => allRecordsSVP.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // ---------------- KPI ----------------

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);

                        break;

                    case 6: // 6 CEEPO
                        $("#reassigned-tab").hide();

                        // Get ALL requests once (for KPIs + all tabs)
                        allRequestsForKPICounts = await GetManagerRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        // KPI – total
                        $(".totalReq").text(myRequestList?.length || 0);

                        // Table handling (same as LM)
                        let targetTableCEEPO = getTableByTabName(elem);
                        let allRecordsCEEPO = [];
                        let loadedBatchesRefCEEPO = { value: 0 };
                        let tableInstanceCEEPO = null;

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsCEEPO = allAllRequests;
                            tableInstanceCEEPO = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefCEEPO = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsCEEPO;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {

                            let ceepoPendingRequests = myRequestList.filter(x =>
                                (x.al_requeststatus === 17 || x.al_requeststatus === 41) ||
                                (x._al_primarycontact_value === loggedInUserId && x.al_requeststatus === 5)
                            );

                            allPendingRequests = [...ceepoPendingRequests];
                            allRecordsCEEPO = allPendingRequests;
                            tableInstanceCEEPO = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefCEEPO = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsCEEPO;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            allReturnedRequests = myRequestList.filter(x => x.al_requeststatus === 5);
                            allRecordsCEEPO = allReturnedRequests;
                            tableInstanceCEEPO = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefCEEPO = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsCEEPO;
                        }

                        // RESOLVED / APPROVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Resolved')) {
                            allResolvedRequests = myRequestList.filter(
                                x => x.al_requeststatus === 6 || x.al_requeststatus === 40
                            );
                            allRecordsCEEPO = allResolvedRequests;
                            tableInstanceCEEPO = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefCEEPO = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsCEEPO;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            allRejectedRequests = myRequestList.filter(x => x.al_requeststatus === 4);
                            allRecordsCEEPO = allRejectedRequests;
                            tableInstanceCEEPO = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefCEEPO = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsCEEPO;
                        }

                        // ---------------- LOAD FIRST BATCH ----------------

                        if (allRecordsCEEPO.length > 0 && tableInstanceCEEPO) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsCEEPO,
                                loadedBatches: loadedBatchesRefCEEPO,
                                tableInstance: tableInstanceCEEPO,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefCEEPO.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('approved') || elem.includes('Resolved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination (same as LM)
                            setTimeout(() => {
                                try {
                                    let api = tableInstanceCEEPO.api();
                                    let settings = api.settings()[0];
                                    settings._iRecordsTotal = allRecordsCEEPO.length;
                                    settings._iRecordsDisplay = allRecordsCEEPO.length;
                                    settings.fnRecordsTotal = () => allRecordsCEEPO.length;
                                    settings.fnRecordsDisplay = () => allRecordsCEEPO.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // ---------------- KPI ----------------

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);

                        break;

                    case 7: // For PSSC Manager
                        $("#reassigned-tab").hide();
                        $("#registerNewReqBtn").hide();
                        $("#grc-approve-tab, #grc-returned-tab, #grc-reject-tab").hide();

                        // Get ALL requests once
                        allRequestsForKPICounts = await GetAllRecordForPSSCManager(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        $(".totalReq").text(myRequestList?.length || 0);

                        let allRecordsPSSCManager = [];
                        let tableInstancePSSCManager = null;
                        let loadedBatchesRefPSSCManager = { value: 0 };

                        // ---------------- TAB LOGIC ----------------

                        // ALL → Re-Assigned
                        if (elem && (elem.includes('All') || elem === 'All')) {
                            allRecordsPSSCManager = myRequestList.filter(x => x.al_requeststatus === 3);
                            tableInstancePSSCManager = allReqTable;
                            allRequestDataStorage = allRecordsPSSCManager;
                        }

                        // PENDING → For Approval
                        else if (elem && (elem.includes('pending') || elem === 'Pending')) {
                            allRecordsPSSCManager = myRequestList.filter(x => x.al_requeststatus === 42);
                            tableInstancePSSCManager = pendingReqTable;
                            allRequestDataStorage = allRecordsPSSCManager;
                            $(".pendingReq").text(allRecordsPSSCManager.length);
                        }

                        // ---------------- LOAD ----------------

                        if (allRecordsPSSCManager.length > 0 && tableInstancePSSCManager) {
                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsPSSCManager,
                                loadedBatches: loadedBatchesRefPSSCManager,
                                tableInstance: tableInstancePSSCManager,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefPSSCManager.value = 1;

                            setTimeout(() => {
                                try {
                                    let api = tableInstancePSSCManager.api();
                                    let s = api.settings()[0];
                                    s._iRecordsTotal = allRecordsPSSCManager.length;
                                    s._iRecordsDisplay = allRecordsPSSCManager.length;
                                    s.fnRecordsTotal = () => allRecordsPSSCManager.length;
                                    s.fnRecordsDisplay = () => allRecordsPSSCManager.length;
                                    api.draw(false);
                                } catch { }
                            }, 100);
                        }

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);
                        break;

                    case 8: // For Approver Team (FSSC Line Manager)
                        $("#reassigned-tab").hide();

                        // Get ALL requests once
                        allRequestsForKPICounts = await GetAllRecordForFSSCManager(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        $(".totalReq").text(myRequestList?.length || 0);

                        let allRecordsFLM = [];
                        let tableInstanceFLM = null;
                        let loadedBatchesRefFLM = { value: 0 };

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'All')) {
                            allRecordsFLM = [...myRequestList];
                            tableInstanceFLM = allReqTable;
                            loadedAllBatches = 0;
                            allRequestDataStorage = allRecordsFLM;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'Pending')) {

                            allRecordsFLM = myRequestList.filter(x =>
                                x.al_requeststatus === 19 ||
                                x.al_requeststatus === 42 ||
                                x.al_requeststatus === 41 ||
                                (x._al_primarycontact_value === loggedInUserId && x.al_requeststatus === 5)
                            );

                            tableInstanceFLM = pendingReqTable;
                            loadedPendingBatches = 0;
                            allRequestDataStorage = allRecordsFLM;
                            $(".pendingReq").text(allRecordsFLM.length);
                        }

                        // RETURNED
                        else if (elem && elem.includes('Returned')) {
                            allRecordsFLM = myRequestList.filter(x => x.al_requeststatus === 5);
                            tableInstanceFLM = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            allRequestDataStorage = allRecordsFLM;
                        }

                        // RESOLVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved'))) {
                            allRecordsFLM = myRequestList.filter(
                                x => x.al_requeststatus === 6 || x.al_requeststatus === 40
                            );
                            tableInstanceFLM = grcApproveTable;
                            loadedResolvedBatches = 0;
                            allRequestDataStorage = allRecordsFLM;
                        }

                        // REJECTED
                        else if (elem && elem.includes('Rejected')) {
                            allRecordsFLM = myRequestList.filter(x => x.al_requeststatus === 4);
                            tableInstanceFLM = grcRejectTable;
                            loadedRejectedBatches = 0;
                            allRequestDataStorage = allRecordsFLM;
                        }

                        // ---------------- LOAD ----------------

                        if (allRecordsFLM.length > 0 && tableInstanceFLM) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsFLM,
                                loadedBatches: loadedBatchesRefFLM,
                                tableInstance: tableInstanceFLM,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefFLM.value = 1;

                            setTimeout(() => {
                                try {
                                    let api = tableInstanceFLM.api();
                                    let s = api.settings()[0];
                                    s._iRecordsTotal = allRecordsFLM.length;
                                    s._iRecordsDisplay = allRecordsFLM.length;
                                    s.fnRecordsTotal = () => allRecordsFLM.length;
                                    s.fnRecordsDisplay = () => allRecordsFLM.length;
                                    api.draw(false);
                                } catch { }
                            }, 100);
                        }

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);
                        break;

                    case 9: // For FSSC Team Member
                        $("#reassigned-tab").hide();
                        $("#assignBtn").removeClass("d-none");

                        // Get ALL requests once (for KPIs + tabs)
                        allRequestsForKPICounts = await GetManagerRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        // KPI – total
                        $(".totalReq").text(myRequestList?.length || 0);

                        let allRecordsFTM = [];
                        let tableInstanceFTM = null;
                        let loadedBatchesRefFTM = { value: 0 };

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsFTM = allAllRequests;
                            tableInstanceFTM = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefFTM = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsFTM;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {

                            let fsscTeamMemberPendingRequests = myRequestList.filter(x =>
                                (x.al_requeststatus === 18 || x.al_requeststatus === 41) ||
                                (x._al_primarycontact_value === loggedInUserId && x.al_requeststatus === 5)
                            );

                            allPendingRequests = [...fsscTeamMemberPendingRequests];
                            allRecordsFTM = allPendingRequests;
                            tableInstanceFTM = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefFTM = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsFTM;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            allReturnedRequests = myRequestList.filter(x => x.al_requeststatus === 5);
                            allRecordsFTM = allReturnedRequests;
                            tableInstanceFTM = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefFTM = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsFTM;
                        }

                        // RESOLVED / APPROVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Resolved')) {
                            allResolvedRequests = myRequestList.filter(
                                x => x.al_requeststatus === 6 || x.al_requeststatus === 40
                            );
                            allRecordsFTM = allResolvedRequests;
                            tableInstanceFTM = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefFTM = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsFTM;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            allRejectedRequests = myRequestList.filter(x => x.al_requeststatus === 4);
                            allRecordsFTM = allRejectedRequests;
                            tableInstanceFTM = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefFTM = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsFTM;
                        }

                        // ---------------- LOAD FIRST BATCH ----------------

                        if (allRecordsFTM.length > 0 && tableInstanceFTM) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsFTM,
                                loadedBatches: loadedBatchesRefFTM,
                                tableInstance: tableInstanceFTM,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefFTM.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('approved') || elem.includes('Resolved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination (same as all other roles)
                            setTimeout(() => {
                                try {
                                    let api = tableInstanceFTM.api();
                                    let s = api.settings()[0];
                                    s._iRecordsTotal = allRecordsFTM.length;
                                    s._iRecordsDisplay = allRecordsFTM.length;
                                    s.fnRecordsTotal = () => allRecordsFTM.length;
                                    s.fnRecordsDisplay = () => allRecordsFTM.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // ---------------- KPI ----------------

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);

                        break;

                    case 10: //P1 Request
                        $("#reassigned-tab").hide();

                        $("#assignBtn").removeClass("d-none");
                        // P1 Request


                        //  let p1Requests = await GetP1Requests(loggedInUserId);
                        //  BindRequest(userWithTeam, p1Requests, p1Table);


                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(myRequestList);
                        break;
                    case 21: // 21 BU Head
                        $('#reassigned-tab').hide();

                        // Get ALL requests once (for KPIs + all tabs)
                        allRequestsForKPICounts = await GetManagerRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        // KPI – total
                        $(".totalReq").text(myRequestList?.length || 0);

                        let allRecordsBUH = [];
                        let tableInstanceBUH = null;
                        let loadedBatchesRefBUH = { value: 0 };

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsBUH = allAllRequests;
                            tableInstanceBUH = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefBUH = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsBUH;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {

                            let buHeadPendingRequests = myRequestList.filter(x =>
                                x.al_requeststatus === 1 ||
                                x.al_requeststatus === 16 ||
                                x.al_requeststatus === 41 ||
                                (x.al_requeststatus === 5 && x._al_primarycontact_value === loggedInUserId) ||
                                x._al_primarycontact_value === loggedInUserId
                            );

                            allPendingRequests = [...buHeadPendingRequests];
                            allRecordsBUH = allPendingRequests;
                            tableInstanceBUH = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefBUH = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsBUH;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            allReturnedRequests = myRequestList.filter(x => x.al_requeststatus === 5);
                            allRecordsBUH = allReturnedRequests;
                            tableInstanceBUH = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefBUH = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsBUH;
                        }

                        // RESOLVED / APPROVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Resolved')) {
                            allResolvedRequests = myRequestList.filter(
                                x => x.al_requeststatus === 6 || x.al_requeststatus === 40
                            );
                            allRecordsBUH = allResolvedRequests;
                            tableInstanceBUH = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefBUH = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsBUH;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            allRejectedRequests = myRequestList.filter(x => x.al_requeststatus === 4);
                            allRecordsBUH = allRejectedRequests;
                            tableInstanceBUH = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefBUH = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsBUH;
                        }

                        // ---------------- LOAD FIRST BATCH ----------------

                        if (allRecordsBUH.length > 0 && tableInstanceBUH) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsBUH,
                                loadedBatches: loadedBatchesRefBUH,
                                tableInstance: tableInstanceBUH,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefBUH.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('approved') || elem.includes('Resolved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination (same pattern everywhere)
                            setTimeout(() => {
                                try {
                                    let api = tableInstanceBUH.api();
                                    let s = api.settings()[0];
                                    s._iRecordsTotal = allRecordsBUH.length;
                                    s._iRecordsDisplay = allRecordsBUH.length;
                                    s.fnRecordsTotal = () => allRecordsBUH.length;
                                    s.fnRecordsDisplay = () => allRecordsBUH.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // ---------------- KPI / NOTIFICATIONS ----------------

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);
                        await getNotificationItems();

                        break;

                    case 22: // 22 Finance Head
                        $('#reassigned-tab').hide();

                        // Get ALL requests once (for KPIs + all tabs)
                        allRequestsForKPICounts = await GetManagerRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        // KPI – total
                        $(".totalReq").text(myRequestList?.length || 0);

                        let allRecordsFH = [];
                        let tableInstanceFH = null;
                        let loadedBatchesRefFH = { value: 0 };

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsFH = allAllRequests;
                            tableInstanceFH = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefFH = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsFH;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {

                            let financeHeadPendingRequests = myRequestList.filter(x =>
                                x.al_requeststatus === 1 ||
                                x.al_requeststatus === 16 ||
                                x.al_requeststatus === 41 ||
                                (x.al_requeststatus === 5 && x._al_primarycontact_value === loggedInUserId) ||
                                x._al_primarycontact_value === loggedInUserId
                            );

                            allPendingRequests = [...financeHeadPendingRequests];
                            allRecordsFH = allPendingRequests;
                            tableInstanceFH = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefFH = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsFH;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            allReturnedRequests = myRequestList.filter(x => x.al_requeststatus === 5);
                            allRecordsFH = allReturnedRequests;
                            tableInstanceFH = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefFH = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsFH;
                        }

                        // RESOLVED / APPROVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Resolved')) {
                            allResolvedRequests = myRequestList.filter(
                                x => x.al_requeststatus === 6 || x.al_requeststatus === 40
                            );
                            allRecordsFH = allResolvedRequests;
                            tableInstanceFH = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefFH = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsFH;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            allRejectedRequests = myRequestList.filter(x => x.al_requeststatus === 4);
                            allRecordsFH = allRejectedRequests;
                            tableInstanceFH = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefFH = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsFH;
                        }

                        // ---------------- LOAD FIRST BATCH ----------------

                        if (allRecordsFH.length > 0 && tableInstanceFH) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsFH,
                                loadedBatches: loadedBatchesRefFH,
                                tableInstance: tableInstanceFH,
                                bindRequestTeamType: -1
                            });

                            loadedBatchesRefFH.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('approved') || elem.includes('Resolved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination
                            setTimeout(() => {
                                try {
                                    let api = tableInstanceFH.api();
                                    let s = api.settings()[0];
                                    s._iRecordsTotal = allRecordsFH.length;
                                    s._iRecordsDisplay = allRecordsFH.length;
                                    s.fnRecordsTotal = () => allRecordsFH.length;
                                    s.fnRecordsDisplay = () => allRecordsFH.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // ---------------- KPI / NOTIFICATIONS ----------------

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);
                        await getNotificationItems();

                        break;

                    case 24: // 24 SVP Procurement
                        $("#reassigned-tab").hide();

                        // Fetch all requests for this user
                        myRequestList = await GetManagerRequests(loggedInUserId);

                        // KPI: total requests
                        $(".totalReq").text(myRequestList?.length || 0);

                        // ---------- ALL TAB ----------
                        allAllRequests = [...myRequestList];
                        loadedAllBatches = 0;
                        await window.loadFirstBatchAndUpdatePagination({
                            allRecords: allAllRequests,
                            loadedBatches: { value: loadedAllBatches },
                            tableInstance: allReqTable,
                            bindRequestTeamType: teamType
                        });
                        loadedAllBatches = 1;
                        setTimeout(() => {
                            if (allReqTable && allAllRequests.length > 0) {
                                window.updateDataTablesPagination(allReqTable, allAllRequests.length);
                            }
                        }, 100);

                        // ---------- PENDING TAB ----------
                        let svpPendingRequests = myRequestList.filter(x => {
                            const approvals = x.al_approvalrequest_Request_al_request;
                            if (!approvals || approvals.length === 0) return false;

                            // get latest approval request by createdon
                            const latestApproval = approvals.reduce((latest, current) =>
                                new Date(current.createdon) > new Date(latest.createdon) ? current : latest
                            );

                            return (
                                x.al_requeststatus === 46 ||
                                x.al_requeststatus === 41 ||
                                (x.al_requeststatus === 49 && latestApproval.al_teamtype === 24 && latestApproval.al_requeststatus === 1) ||
                                (x._al_primarycontact_value === loggedInUserId && x.al_requeststatus === 5)
                            );
                        });

                        allPendingRequests = [...svpPendingRequests];
                        loadedPendingBatches = 0;
                        await window.loadFirstBatchAndUpdatePagination({
                            allRecords: allPendingRequests,
                            loadedBatches: { value: loadedPendingBatches },
                            tableInstance: pendingReqTable,
                            bindRequestTeamType: -1
                        });
                        loadedPendingBatches = 1;
                        $(".pendingReq").text(allPendingRequests.length);
                        setTimeout(() => {
                            if (pendingReqTable && allPendingRequests.length > 0) {
                                window.updateDataTablesPagination(pendingReqTable, allPendingRequests.length);
                            }
                        }, 100);

                        // ---------- RETURNED TAB ----------
                        allReturnedRequests = myRequestList.filter(x => x.al_requeststatus === 5);
                        loadedReturnedBatches = 0;
                        await window.loadFirstBatchAndUpdatePagination({
                            allRecords: allReturnedRequests,
                            loadedBatches: { value: loadedReturnedBatches },
                            tableInstance: returenedRequestTable,
                            bindRequestTeamType: teamType
                        });
                        loadedReturnedBatches = 1;
                        setTimeout(() => {
                            if (returenedRequestTable && allReturnedRequests.length > 0) {
                                window.updateDataTablesPagination(returenedRequestTable, allReturnedRequests.length);
                            }
                        }, 100);

                        // ---------- RESOLVED TAB ----------
                        allResolvedRequests = myRequestList.filter(x => x.al_requeststatus === 6 || x.al_requeststatus === 40);
                        loadedResolvedBatches = 0;
                        await window.loadFirstBatchAndUpdatePagination({
                            allRecords: allResolvedRequests,
                            loadedBatches: { value: loadedResolvedBatches },
                            tableInstance: grcApproveTable,
                            bindRequestTeamType: teamType
                        });
                        loadedResolvedBatches = 1;
                        setTimeout(() => {
                            if (grcApproveTable && allResolvedRequests.length > 0) {
                                window.updateDataTablesPagination(grcApproveTable, allResolvedRequests.length);
                            }
                        }, 100);

                        // ---------- REJECTED TAB ----------
                        allRejectedRequests = myRequestList.filter(x => x.al_requeststatus === 4);
                        loadedRejectedBatches = 0;
                        await window.loadFirstBatchAndUpdatePagination({
                            allRecords: allRejectedRequests,
                            loadedBatches: { value: loadedRejectedBatches },
                            tableInstance: grcRejectTable,
                            bindRequestTeamType: teamType
                        });
                        loadedRejectedBatches = 1;
                        setTimeout(() => {
                            if (grcRejectTable && allRejectedRequests.length > 0) {
                                window.updateDataTablesPagination(grcRejectTable, allRejectedRequests.length);
                            }
                        }, 100);

                        // Show KPI cards
                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(myRequestList);

                        break;

                    case 25: // 25 Finance
                        $('#reassigned-tab').hide();

                        // Get ALL requests once (for KPIs + tabs)
                        allRequestsForKPICounts = await GetManagerRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        let allRecordsFinance = [];
                        let tableInstanceFinance = null;
                        let loadedBatchesRefFinance = { value: 0 };

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsFinance = allAllRequests;
                            tableInstanceFinance = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefFinance = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsFinance;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {

                            let financePendingRequests = myRequestList.filter(x => {
                                if (x.al_requeststatus == 1 || x.al_requeststatus == 16 || x.al_requeststatus == 41) {
                                    return true;
                                } else if (x.al_requeststatus == 5 && x._al_primarycontact_value == loggedInUserId) {
                                    return true;
                                } else if (x._al_primarycontact_value == loggedInUserId) {
                                    return true;
                                } else {
                                    return false;
                                }
                            });

                            allPendingRequests = [...financePendingRequests];
                            allRecordsFinance = allPendingRequests;
                            tableInstanceFinance = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefFinance = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsFinance;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            let returnedRequestListFinance = myRequestList.filter(x => x.al_requeststatus == 5);
                            allReturnedRequests = [...returnedRequestListFinance];
                            allRecordsFinance = allReturnedRequests;
                            tableInstanceFinance = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefFinance = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsFinance;
                        }

                        // RESOLVED / APPROVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Resolved')) {
                            let approvedRequestListFinance = myRequestList.filter(x => x.al_requeststatus == 6 || x.al_requeststatus == 40);
                            allResolvedRequests = [...approvedRequestListFinance];
                            allRecordsFinance = allResolvedRequests;
                            tableInstanceFinance = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefFinance = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsFinance;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            let rejectedRequestListFinance = myRequestList.filter(x => x.al_requeststatus == 4);
                            allRejectedRequests = [...rejectedRequestListFinance];
                            allRecordsFinance = allRejectedRequests;
                            tableInstanceFinance = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefFinance = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsFinance;
                        }

                        // ---------------- LOAD FIRST BATCH ----------------

                        if (allRecordsFinance.length > 0 && tableInstanceFinance) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsFinance,
                                loadedBatches: loadedBatchesRefFinance,
                                tableInstance: tableInstanceFinance,
                                bindRequestTeamType: -1
                            });

                            loadedBatchesRefFinance.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('approved') || elem.includes('Resolved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination
                            setTimeout(() => {
                                try {
                                    let api = tableInstanceFinance.api();
                                    let s = api.settings()[0];
                                    s._iRecordsTotal = allRecordsFinance.length;
                                    s._iRecordsDisplay = allRecordsFinance.length;
                                    s.fnRecordsTotal = () => allRecordsFinance.length;
                                    s.fnRecordsDisplay = () => allRecordsFinance.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // ---------------- KPI ----------------

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);
                        await getNotificationItems();

                        break;

                    case 23: // 23 Category Manager
                        $('#reassigned-tab').hide();

                        // Get ALL requests once (for KPIs + tabs)
                        allRequestsForKPICounts = await GetManagerRequests(loggedInUserId);
                        myRequestList = allRequestsForKPICounts;

                        let allRecordsCM = [];
                        let tableInstanceCM = null;
                        let loadedBatchesRefCM = { value: 0 };

                        // ---------------- TAB LOGIC ----------------

                        // ALL
                        if (elem && (elem.includes('All') || elem === 'all requests' || elem === 'All')) {
                            allAllRequests = [...myRequestList];
                            allRecordsCM = allAllRequests;
                            tableInstanceCM = allReqTable;
                            loadedAllBatches = 0;
                            loadedBatchesRefCM = { value: loadedAllBatches };
                            allRequestDataStorage = allRecordsCM;
                        }

                        // PENDING
                        else if (elem && (elem.includes('pending') || elem === 'pending requests' || elem === 'Pending')) {
                            let categoryManagerPendingRequests = myRequestList.filter(x =>
                                x.al_requeststatus == 1 ||
                                x.al_requeststatus == 16 ||
                                x.al_requeststatus == 41 ||
                                (x.al_requeststatus == 5 && x._al_primarycontact_value == loggedInUserId) ||
                                x._al_primarycontact_value == loggedInUserId
                            );

                            allPendingRequests = [...categoryManagerPendingRequests];
                            allRecordsCM = allPendingRequests;
                            tableInstanceCM = pendingReqTable;
                            loadedPendingBatches = 0;
                            loadedBatchesRefCM = { value: loadedPendingBatches };
                            allRequestDataStorage = allRecordsCM;

                            $(".pendingReq").text(allPendingRequests.length);
                        }

                        // RETURNED
                        else if (elem && (elem.includes('Returned') || elem === 'returned requests' || elem === 'Returned')) {
                            let returnedRequestListCategoryManager = myRequestList.filter(x => x.al_requeststatus == 5);
                            allReturnedRequests = [...returnedRequestListCategoryManager];
                            allRecordsCM = allReturnedRequests;
                            tableInstanceCM = returenedRequestTable;
                            loadedReturnedBatches = 0;
                            loadedBatchesRefCM = { value: loadedReturnedBatches };
                            allRequestDataStorage = allRecordsCM;
                        }

                        // RESOLVED / APPROVED
                        else if (elem && (elem.includes('approved') || elem.includes('Resolved') || elem === 'Resolved')) {
                            let approvedRequestListCategoryManager = myRequestList.filter(x => x.al_requeststatus == 6 || x.al_requeststatus == 40);
                            allResolvedRequests = [...approvedRequestListCategoryManager];
                            allRecordsCM = allResolvedRequests;
                            tableInstanceCM = grcApproveTable;
                            loadedResolvedBatches = 0;
                            loadedBatchesRefCM = { value: loadedResolvedBatches };
                            allRequestDataStorage = allRecordsCM;
                        }

                        // REJECTED
                        else if (elem && (elem.includes('Rejected') || elem === 'rejected requests' || elem === 'Rejected')) {
                            let rejectedRequestListCategoryManager = myRequestList.filter(x => x.al_requeststatus == 4);
                            allRejectedRequests = [...rejectedRequestListCategoryManager];
                            allRecordsCM = allRejectedRequests;
                            tableInstanceCM = grcRejectTable;
                            loadedRejectedBatches = 0;
                            loadedBatchesRefCM = { value: loadedRejectedBatches };
                            allRequestDataStorage = allRecordsCM;
                        }

                        // ---------------- LOAD FIRST BATCH ----------------

                        if (allRecordsCM.length > 0 && tableInstanceCM) {

                            await window.loadFirstBatchAndUpdatePagination({
                                allRecords: allRecordsCM,
                                loadedBatches: loadedBatchesRefCM,
                                tableInstance: tableInstanceCM,
                                bindRequestTeamType: teamType
                            });

                            loadedBatchesRefCM.value = 1;

                            if (elem.includes('All')) loadedAllBatches = 1;
                            else if (elem.includes('pending')) loadedPendingBatches = 1;
                            else if (elem.includes('Returned')) loadedReturnedBatches = 1;
                            else if (elem.includes('approved') || elem.includes('Resolved')) loadedResolvedBatches = 1;
                            else if (elem.includes('Rejected')) loadedRejectedBatches = 1;

                            // Fix DataTables pagination
                            setTimeout(() => {
                                try {
                                    let api = tableInstanceCM.api();
                                    let s = api.settings()[0];
                                    s._iRecordsTotal = allRecordsCM.length;
                                    s._iRecordsDisplay = allRecordsCM.length;
                                    s.fnRecordsTotal = () => allRecordsCM.length;
                                    s.fnRecordsDisplay = () => allRecordsCM.length;
                                    api.draw(false);
                                } catch (e) {
                                    console.warn("Pagination update failed", e);
                                }
                            }, 100);
                        }

                        // ---------------- KPI ----------------

                        $(".kpiCards").removeClass("d-none");
                        getRequestCount(allRequestsForKPICounts);
                        await getNotificationItems();

                        break;

                    case -1:
                        showErrorModal("Error fetching team details.");
                        console.log("Error fetching team");
                        break;
                    default:
                        break;
                }
            } catch (error) {
                console.error("Error in switch statement:", error);
                showErrorModal("Error loading request data. Please refresh the page.");
            }
            getFilterOptions();
            $('.custom-loader').css("display", "none");
        } catch (error) {
            console.error("Error in onLoad:", error);
            showErrorModal("Error loading data. Please refresh the page.");
            $('.custom-loader').css("display", "none");
        }
    } else {
        $('.custom-loader').css("display", "none");
    }
    // $('#requestSubtype').select2();
    $('#date').select2();
    $("#dateFilter").select2();
    // $('#legalEntityFilter').select2();
    // $('#assignToFilter').select2();
    // $('#approvalStatus').select2();
    // $('#finalStatus').select2();
    $("#totalReq").select2();
    $("#pendingReq").select2();
    $("#returnReq").select2();
    $("#completedReq").select2();
    // $('#requestSubtype').select2({
    //     placeholder: "Select request sub types",
    //     allowClear: true,
    //     width: '100%',
    //     templateResult: formatOption,
    //     templateSelection: formatOptionSelection
    // });

    // Update checkbox state when selection changes
    // $('#requestSubtype').on('change', function () {
    //     $('#requestSubtype option').each(function () {
    //         $(this).prop('selected', $('#requestSubtype').val() && $('#requestSubtype').val().includes($(this).val()));
    //     });
    // });

    initializeSelect2WithCount('#requestSubtype', 'Sub Request Type');
    initializeSelect2WithCount('#legalEntityFilter', 'Legal Entity');
    initializeSelect2WithCount('#assignToFilter', 'Assign To');
    initializeSelect2WithCount('#approvalStatus', 'Approval Status');
    initializeSelect2WithCount('#finalStatus', 'Final Status');
}
// function formatOption(option) {
//     if (!option.id) return option.text;
//     var $option = $(
//         '<span><input type="checkbox" class="option-checkbox subReqTypeCheckBox" ' + (option.selected ? 'checked' : '') + ' /> ' + option.text + '</span >'
//     );
//     return $option;
// }
// let isExist = true;
// function formatOptionSelection(option) {
//     if (!option.id) return option.text;
//     var selectedOptions = $('#requestSubtype').val() || [];
//     // if (selectedOptions.length > 1 && !isExist) {
//     //     return;
//     // }
//     if (selectedOptions.length == 0) {
//         return "Select request sub types"
//     } else if (selectedOptions.length == 1) {
//         return $('#requestSubtype :selected').text();
//     } else if (selectedOptions.length > 1 && isExist){
//         isExist = false;
//         return "Multiple Selected";
//     }
// }
// // Update checkbox state when selection changes
// $('#requestSubtype').on('change', function () {
//     $('#requestSubtype option').each(function () {
//         $(this).prop('selected', $('#requestSubtype').val() && $('#requestSubtype').val().includes($(this).val()));
//     });
// });

async function GetPendingData(userId) {
    let todayDateTime = new Date().toISOString();
    let reqList = await GetEntityList(
        "al_requests",
        `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime}`,
        "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendorname,al_requeststatus,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_reassigndetail,al_subrequesttype,al_vendormultiplelabel,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority",
        "modifiedon desc",
        "al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon),al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)"
    );
    return reqList;
}


function initializeSelect2WithCount(id, placeholder) {
    const $select = $(id);

    $select.select2({
        placeholder: placeholder,
        allowClear: true,
        closeOnSelect: false,
        width: '100%',
        minimumResultsForSearch: 0, // ✅ Always show search bar
        templateResult: function (data) {
            if (!data.id) return data.text;
            const selected = $select.val() || [];
            const isChecked = selected.includes(data.id) ? 'checked' : '';
            return $(`<div class="d-flex align-items-center">
<input type="checkbox" class="form-check-input" style="margin-right: 8px;" ${isChecked} />
<span>${data.text}</span>
</div>`);
        },
        templateSelection: function (data) {
            let checkData = data;
            console.log("temp slectiondata===>", checkData);
            const count = data.length;
            return count === 0 ? "" : `${count} selected`;
        },
        escapeMarkup: function (markup) {
            return markup;
        }
    });

    // Force re-render count on change
    $select.on('change', function () {
        const selectedData = $select.select2('data');
        const count = selectedData.length;
        var label = $select.siblings('label').text();
        $(`${id}`).next('.select2-container').find('.select2-selection__rendered')
            .html(count === 0 ? `${label}` : `${count} selected`);
    });

    // Optional: Keep dropdown open on click
    $select.on('select2:select select2:unselect', function () {
        setTimeout(() => {
            $select.select2('close').select2('open');
        }, 0);
    });
}


async function BindRequest(userWithTeam, reqList, tableElement, teamType = -1, skip = null, top = null, append = false, skipKPIUpdate = false) {
    pendingForCount = reqList;
    // $(".pendingReq").text(reqList?.length);

    // Detect if this is the pending table - must be defined at function start for all scopes
    let isPendingTableForAll = tableElement.table().node().id === 'pendingReqTable';

    // Pre-fetch SLA settings for performance (don't let errors here stop data loading)
    try {
        await getSLASettings();
    } catch (error) {
        console.warn("SLA settings pre-fetch failed, continuing without SLA data:", error);
    }

    // Helper function to ensure calculation variables are always defined
    async function getColumnData(result) {
        let assignedDateTime = getAssignedDateTime(result);
        let dueDate = await calculateDueDate(result);
        return {
            assignedDateTimeHtml: assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted,
            hoursRemainingHtml: dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted
        };
    }

    if (teamType == -1) {  //For All Pending Request
        if (!append) {
            tableElement.clear();
        } else {
            // When appending, we'll save page after draw (see below)
        }
        // Update pending count with total records if available (skip if clearing filters)
        if (!skipKPIUpdate) {
            if (allPendingRequests.length > 0) {
                $(".pendingReq").text(allPendingRequests.length);
            } else {
                $(".pendingReq").text(reqList?.length);
            }
        }

        // Detect actual column positions in HTML table (for non-pending tables)
        let assignedDateIdx = 8, dueDateIdx = 9, approvalStatusIdx = 10, finalStatusIdx = 11; // Default positions

        if (!isPendingTableForAll) {
            try {
                let $table = $(tableElement.table().node());
                let $headers = $table.find('thead tr th');

                $headers.each(function (index) {
                    let headerText = $(this).text().trim().toLowerCase();
                    if (headerText.includes('assigned date') || headerText.includes('assined date')) {
                        assignedDateIdx = index;
                    } else if (headerText.includes('due date')) {
                        dueDateIdx = index;
                    } else if (headerText.includes('approval status')) {
                        approvalStatusIdx = index;
                    } else if (headerText.includes('final status')) {
                        finalStatusIdx = index;
                    }
                });
            } catch (e) {
                console.log("Error detecting column positions:", e);
            }
        }

        for (const result of reqList) {
            try {

                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let requestId = result["al_requestid"]; // Guid
                let subReqType = result.al_subrequesttype;

                //let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let eAuctionName = "N/A";
                let vendorName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                }
                else if (result["al_eauctionname"]) {
                    vendorName = "N/A"
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let tooltipText = `PO Number: ${poNumber || 'N/A'}\nPR Number: ${prNumber || 'N/A'}\nVendor Name: ${vendorName || 'N/A'}`;
                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] : "N/A";

                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                //change
                let psscReAssignBtn = "";
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");
                let firstColumn;
                // Check if user can review this request (for All, Returned, Resolved, Rejected tabs)
                let reqStatus = result.al_requeststatus;
                const canReview =
                    reqStatus == 1 ||
                    (reqStatus == 5 && result._al_primarycontact_value == loggedInUserId) ||
                    reqStatus == 16 ||
                    reqStatus == 41 ||
                    reqStatus == 2 ||
                    reqStatus == 18 ||
                    (reqStatus == 15 && subReqType != 8 && subReqType == 4);

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                // Initialize with default values to prevent undefined errors
                let assignedDateTimeHtml = "-";
                let hoursRemainingHtml = "-";
                try {
                    let assignedDateTime = getAssignedDateTime(result);
                    let dueDate = await calculateDueDate(result);
                    assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                    hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;
                } catch (calcError) {
                    console.warn("Error calculating assigned date/time or due date:", calcError);
                    // Variables already initialized with "-" above
                }
                //PSSC MANAGER and pssc member
                if (userWithTeam.al_teamtype == 3 && result.al_psscreassignmemberid == null && userWithTeam._al_psscmanager_value && result.al_subrequesttype !== 13 && result.al_subrequesttype !== 14) {//PSSC In Pending Tab
                    psscReAssignBtn = `
                        <button type="button" class="btn btn-sm btn-primary ml-1" onclick="openPsscAssignModal(
                        '${userWithTeam._al_psscmanager_value}','${requestId}',
                        '${result._al_currentapprovalrequest_value}',
                        '${result["al_referencenumber"]}',
                        '${result["al_vendorname"]}',
                        '${vendorCode}',
                        '${vendorEmail}',
                        '${result["al_legalentitydetails"]}',
                        '${assignedTo}',
                        '${result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["createdon@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_priority@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_changetype@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_requeststatus@OData.Community.Display.V1.FormattedValue"]}'
                        )">Assign</button>
                    `;
                } else if (userWithTeam.al_teamtype == 7 && result.al_requeststatus == 42 && result.al_psscreassignmemberid) { //PSSC Manager Take Action
                    psscReAssignBtn = `<button type="button" class="btn btn-sm btn-secondary mr-1" 
                         onclick="openPsscModal(
                        '${result.al_psscreassignmemberid}',
                        '${result.al_psscreassignmembername}',
                        '${requestId}',
                        '${result._al_currentapprovalrequest_value}',
                        '${result["al_referencenumber"]}',
                        '${result["al_vendorname"]}',
                        '${vendorCode}',
                        '${vendorEmail}',
                        '${result["al_legalentitydetails"]}',
                        '${assignedTo}',
                        '${result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["createdon@OData.Community.Display.V1.FormattedValue"]}'
                        )">Take Action</button>`;
                    $('.selectioncolumn').removeClass('d-none'); // Still keep the column visible for alignment
                    firstColumn = "";
                    // 0 <th>checkbox column</th>
                    // 1 <th>Request ID</th>  referenceNumber
                    // 2 <th>UVP Request ID</th> uvpRequestId
                    // 3 <th>Request Sub Type</th> subrequesttypeFormatted
                    // 4 <th>Date & Time</th> createdOn
                    // 5 <th>Vendor Name</th> vendorName
                    // 6 <th>Legal Entity</th> legalEntityLabel
                    // 7 <th>Assigned To</th> assignedTo
                    // 8 <th>Requested By</th> requestedBy
                    // 9 <th>Approval Status</th> requestStatus
                    // 10 <th>Final Status</th> finalRequestStatusFormatted
                    // 11 <th>Action</th>
                    // 12 <th>PR Number</th> prNumber
                    // 13 <th>PO Number</th> poNumber

                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: uvpRequestId, 3: subrequesttypeFormatted, 4: createdOn, 5: `<span>${vendorName}</span>`, 6: `<span>${legalEntityLabel}</span>`, 7: assignedTo, 8: requestedBy, 9: requestStatus, 10: finalRequestStatusFormatted, 11: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`, 12: prNumber, 13: poNumber,
                    //     14: ProjectID, 15: EAuctionName
                    // });
                    // Check if table has checkbox (pendingReqTable) or not (allReqTable, etc.)
                    // Use the variable declared at the start of the function
                    if (isPendingTableForAll) {
                        // pendingReqTable has checkbox at index 0
                        // HTML columns: 0=checkbox, 1=RequestID, 2=UVPRequestID, 3=RequestSubType, 4=Date&Time, 5=VendorName, 6=VendorCode, 7=VendorEmail, 8=LegalEntity, 9=AssignedTo, 10=RequestedBy, 11=AssignedDate, 12=DueDate, 13=ApprovalStatus, 14=FinalStatus, 15=Action, 16=PRNumber, 17=PONumber
                        tableElement.row.add({
                            DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: uvpRequestId, 3: subrequesttypeFormatted, 4: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 5: `<span>${vendorName}</span>`, 6: vendorCode, 7: vendorEmail, 8: `<span>${legalEntityLabel}</span>`, 9: assignedTo, 10: requestedBy, 11: assignedDateTimeHtml, 12: hoursRemainingHtml, 13: requestStatus, 14: finalRequestStatusFormatted, 15: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`, 16: prNumber, 17: poNumber,
                            18: ProjectID, 19: "", 20: EAuctionName, 21: "", 22: ""
                        });
                        tableElement.column(0).visible(false);
                    } else {
                        // allReqTable and other tables don't have checkbox - start with referenceNumber at index 0
                        tableElement.row.add({
                            DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result,
                            0: referenceNumber,
                            1: uvpRequestId,
                            2: subrequesttypeFormatted,
                            3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                            4: `<span>${vendorName}</span>`,
                            5: vendorCode,
                            6: vendorEmail,
                            7: `<span>${legalEntityLabel}</span>`,
                            8: assignedTo,
                            9: requestedBy,
                            10: assignedDateTimeHtml,
                            11: hoursRemainingHtml,
                            12: requestStatus,
                            13: finalRequestStatusFormatted,
                            14: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`,
                            15: prNumber,
                            16: poNumber,
                            17: ProjectID,
                            18: "",
                            19: EAuctionName,
                            20: "",
                            21: ""
                        });
                    }
                }
                // Show all columns for all tabs and all roles
                if (isPendingTableForAll) {
                    // For pendingReqTable: Show all columns
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                    tableElement.column(10).visible(true); // Approval Status
                    tableElement.column(11).visible(true); // Final Status
                    tableElement.column(12).visible(true); // Action
                    tableElement.column(13).visible(true); // Actionn
                    tableElement.column(14).visible(true); // Action
                } else {
                    // For other tables: Show all columns
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }


                //For Matrix Approval Request (FSSC MANAGER and FSSC MEMBER)
                if (userWithTeam.al_teamtype == 9 && result.al_psscreassignmemberid == null && userWithTeam._al_fsscmanager_value && result.al_subrequesttype !== 13 && result.al_subrequesttype !== 14) {//FSSC Member In Pending Tab
                    psscReAssignBtn = `
                        <button type="button" class="btn btn-sm btn-primary ml-1" onclick="openFSSCAssignModal(
                        '${userWithTeam._al_fsscmanager_value}',
                        '${requestId}',
                        '${result._al_currentapprovalrequest_value}',
                        '${result["al_referencenumber"]}',
                        '${result["al_vendorname"]}',
                        '${result["al_legalentitydetails"]}',
                        '${assignedTo}',
                        '${vendorCode}',
                        '${vendorEmail}',
                        '${result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["createdon@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_priority@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_changetype@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_requeststatus@OData.Community.Display.V1.FormattedValue"]}'
                        )">Assign</button>
                    `;
                } else if (userWithTeam.al_teamtype == 8 && result.al_requeststatus == 42 && result.al_psscreassignmemberid) {  //FSSC Line Manager Take Action
                    psscReAssignBtn = `<button type="button" class="btn btn-sm btn-secondary mr-1" 
                         onclick="openPsscModal(
                        '${result.al_psscreassignmemberid}',
                        '${result.al_psscreassignmembername}',
                        '${requestId}',
                        '${result._al_currentapprovalrequest_value}',
                        '${result["al_referencenumber"]}',
                        '${result["al_vendorname"]}',
                        '${result["al_legalentitydetails"]}',
                        '${assignedTo}',
                        '${vendorCode}',
                        '${vendorEmail}',
                        '${result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["createdon@OData.Community.Display.V1.FormattedValue"]}'
                        )">Take Action
                        </button>`;
                    // Check if table has checkbox (pendingReqTable) or not (allReqTable, etc.)
                    // Use the variable declared at the start of the function
                    if (isPendingTableForAll) {
                        $('.selectioncolumn').removeClass('d-none'); // Still keep the column visible for alignment
                        firstColumn = "";
                    } else {
                        firstColumn = "";
                    }

                    if (isPendingTableForAll) {
                        // pendingReqTable has checkbox at index 0
                        // HTML columns: 0=checkbox, 1=RequestID, 2=UVPRequestID, 3=RequestSubType, 4=Date&Time, 5=VendorName, 6=VendorCode, 7=VendorEmail, 8=LegalEntity, 9=AssignedTo, 10=RequestedBy, 11=AssignedDate, 12=DueDate, 13=ApprovalStatus, 14=FinalStatus, 15=Action, 16=PRNumber, 17=PONumber
                        tableElement.row.add({
                            DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: uvpRequestId, 3: subrequesttypeFormatted, 4: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 5: `<span>${vendorName}</span>`, 6: vendorCode, 7: vendorEmail, 8: `<span>${legalEntityLabel}</span>`, 9: assignedTo, 10: requestedBy, 11: assignedDateTimeHtml, 12: hoursRemainingHtml, 13: requestStatus, 14: finalRequestStatusFormatted, 15: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`, 16: prNumber, 17: poNumber,
                            18: ProjectID, 19: "", 20: EAuctionName, 21: "", 22: ""
                        });
                    } else {
                        // allReqTable and other tables don't have checkbox - start with referenceNumber at index 0
                        tableElement.row.add({
                            DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result,
                            0: referenceNumber,
                            1: uvpRequestId,
                            2: subrequesttypeFormatted,
                            3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                            4: `<span>${vendorName}</span>`,
                            5: vendorCode,
                            6: vendorEmail,
                            7: `<span>${legalEntityLabel}</span>`,
                            8: assignedTo,
                            9: requestedBy,
                            10: assignedDateTimeHtml,
                            11: hoursRemainingHtml,
                            12: requestStatus,
                            13: finalRequestStatusFormatted,
                            14: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`,
                            15: prNumber,
                            16: poNumber,
                            17: ProjectID,
                            18: "",
                            19: EAuctionName,
                            20: "",
                            21: ""
                        });
                    }
                    // tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: uvpRequestId, 3: requesttypeFormatted, 4: subrequesttypeFormatted, 5: legalEntityLabel, 6: vendorName, 7: prNumber, 8: poNumber, 9: createdOn, 10: assignedTo, 11: requestStatus, 12: finalRequestStatusFormatted,
                    //     13: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`, 14: ProjectID, 15: EAuctionName
                    // });
                    tableElement.column(0).visible(false);
                } else if (userWithTeam.al_teamtype == 8 && (result.al_requeststatus == 19 || result.al_requeststatus == 41)) {
                    // Build action button based on canReview condition
                    let reviewActionHtml = "";
                    if (canReview) {
                        reviewActionHtml = `<a href="/request-journey/?id=${requestId}" title="Review" class="btn btn-sm btn-secondary mr-1">Review</a>`;
                    }
                    psscReAssignBtn = reviewActionHtml;
                    // Check if table has checkbox (pendingReqTable) or not (allReqTable, etc.)
                    // Use the variable declared at the start of the function
                    if (isPendingTableForAll) {
                        $('.selectioncolumn').removeClass('d-none'); // Still keep the column visible for alignment
                        firstColumn = "";
                    } else {
                        firstColumn = "";
                    }

                    if (isPendingTableForAll) {
                        // pendingReqTable has checkbox at index 0
                        // HTML columns: 0=checkbox, 1=RequestID, 2=UVPRequestID, 3=RequestSubType, 4=Date&Time, 5=VendorName, 6=VendorCode, 7=VendorEmail, 8=LegalEntity, 9=AssignedTo, 10=RequestedBy, 11=AssignedDate, 12=DueDate, 13=ApprovalStatus, 14=FinalStatus, 15=Action, 16=PRNumber, 17=PONumber
                        tableElement.row.add({
                            DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                            __raw: result,
                            0: firstColumn,
                            1: referenceNumber,
                            2: uvpRequestId,
                            3: subrequesttypeFormatted,
                            4: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                            5: `<span>${vendorName}</span>`,
                            6: vendorCode,
                            7: vendorEmail,
                            8: `<span>${legalEntityLabel}</span>`,
                            9: assignedTo,
                            10: requestedBy,
                            11: assignedDateTimeHtml,
                            12: hoursRemainingHtml,
                            13: requestStatus,
                            14: finalRequestStatusFormatted,
                            15: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`,
                            16: prNumber,
                            17: poNumber,
                            18: ProjectID,
                            19: "",
                            20: EAuctionName,
                            21: "",
                            22: ""
                        });
                    } else {
                        // allReqTable and other tables don't have checkbox - start with referenceNumber at index 0
                        tableElement.row.add({
                            DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                            __raw: result,
                            0: referenceNumber,
                            1: uvpRequestId,
                            2: subrequesttypeFormatted,
                            3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                            4: `<span>${vendorName}</span>`,
                            5: vendorCode,
                            6: vendorEmail,
                            7: `<span>${legalEntityLabel}</span>`,
                            8: assignedTo,
                            9: requestedBy,
                            10: assignedDateTimeHtml,
                            11: hoursRemainingHtml,
                            12: requestStatus,
                            13: finalRequestStatusFormatted,
                            14: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`,
                            15: prNumber,
                            16: poNumber,
                            17: ProjectID,
                            18: "",
                            19: EAuctionName,
                            20: "",
                            21: "",
                            22: ""
                        });
                    }
                    // tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: uvpRequestId, 3: requesttypeFormatted, 4: subrequesttypeFormatted, 5: legalEntityLabel, 6: vendorName, 7: prNumber, 8: poNumber, 9: createdOn, 10: assignedTo, 11: requestStatus, 12: finalRequestStatusFormatted,
                    //     13: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`, 14: ProjectID, 15: EAuctionName
                    // });
                    tableElement.column(0).visible(false);
                }
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                } else {
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                }
                // OTHER TEAMS
                if (userWithTeam.al_teamtype != 7 && userWithTeam.al_teamtype != 8) {
                    // Check if table has checkbox (pendingReqTable) or not (allReqTable, etc.)
                    // Use the variable declared at the start of the function
                    if (isPendingTableForAll) {
                        if (subReqType != 13 && subReqType != 14) {
                            firstColumn = `<div class="check-box">
                        <input type="checkbox" class="form-check-input reassignChkBox" id="${requestId}" data-id="${result._al_currentapprovalrequest_value}">
                        <label class="form-check-label" for="${requestId}"></label>
                        </div>`;
                        } else {
                            firstColumn = `<div class="check-box d-none">
                        <input type="checkbox" class="form-check-input reassignChkBox d-none" id="${requestId}" data-id="${requestId}" disabled>
                        <label class="form-check-label" for="${requestId}"></label>
                        </div>`;
                        }
                    } else {
                        // For tables without checkbox, don't use firstColumn
                        firstColumn = "";
                    }

                    if ((userWithTeam.al_teamtype == 3 || userWithTeam.al_teamtype == 9)) {//&& subReqType != 13 && subReqType != 14

                        if (isPendingTableForAll) {
                            $('.selectioncolumn').removeClass('d-none');
                        }
                        // Build action button based on canReview condition

                        let reviewActionHtml = "";

                        if (canReview) {

                            reviewActionHtml = `<a href="/request-journey/?id=${requestId}" title="Review" class="btn btn-sm btn-secondary mr-1">Review</a>`;

                        }

                        let actionHtml = `<div class="d-flex justify-content-center">${reviewActionHtml}${psscReAssignBtn}</div>`;

                        if (isPendingTableForAll) {
                            // pendingReqTable has checkbox at index 0
                            // HTML columns: 0=checkbox, 1=RequestID, 2=UVPRequestID, 3=RequestSubType, 4=Date&Time, 5=VendorName, 6=VendorCode, 7=VendorEmail, 8=LegalEntity, 9=AssignedTo, 10=RequestedBy, 11=AssignedDate, 12=DueDate, 13=ApprovalStatus, 14=FinalStatus, 15=Action, 16=PRNumber, 17=PONumber
                            tableElement.row.add({
                                DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: uvpRequestId, 3: subrequesttypeFormatted, 4: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 5: `<span>${vendorName}</span>`, 6: vendorCode, 7: vendorEmail, 8: `<span>${legalEntityLabel}</span>`, 9: assignedTo, 10: requestedBy, 11: assignedDateTimeHtml, 12: hoursRemainingHtml, 13: requestStatus, 14: finalRequestStatusFormatted, 15: actionHtml, 16: prNumber, 17: poNumber,
                                18: ProjectID, 19: "", 20: EAuctionName, 21: "", 22: ""
                            });
                        } else {
                            // allReqTable and other tables don't have checkbox - start with referenceNumber at index 0
                            tableElement.row.add({
                                DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: actionHtml, 15: prNumber, 16: poNumber,
                                17: ProjectID, 18: "", 19: EAuctionName, 20: "", 21: ""
                            });
                        }
                        // tableElement.row.add({
                        //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: uvpRequestId, 3: requesttypeFormatted, 4: subrequesttypeFormatted, 5: legalEntityLabel, 6: vendorName, 7: prNumber, 8: poNumber, 9: createdOn, 10: assignedTo, 11: requestStatus, 12: finalRequestStatusFormatted, 13: `<div class="d-flex justify-content-center"><a href="/register-new-request/?id=` + requestId + `" title="Review" class="btn btn-sm btn-secondary mr-1">Review</a>${psscReAssignBtn}</div>`, 14: ProjectID, 15: EAuctionName
                        // });
                    } else {
                        if (isPendingTableForAll) {
                            $('.selectioncolumn').removeClass('d-none');

                        }
                        // Build action button based on canReview condition

                        let reviewActionHtml = "";

                        if (canReview) {

                            reviewActionHtml = `<a href="/request-journey/?id=${requestId}" title="Review" class="btn btn-sm btn-secondary mr-1">Review</a>`;

                        }

                        let actionHtmlForPending = `<div class="d-flex justify-content-center">${reviewActionHtml}${psscReAssignBtn}</div>`;

                        if (isPendingTableForAll) {
                            // pendingReqTable has checkbox at index 0
                            // Column structure: 0=checkbox(hidden), 1=RequestID, 2=UVPRequestID, 3=RequestSubType, 4=Date&Time, 5=VendorName, 6=VendorCode, 7=VendorEmail, 8=LegalEntity, 9=AssignedTo, 10=RequestedBy, 11=AssignedDate, 12=DueDate, 13=ApprovalStatus, 14=FinalStatus, 15=Action
                            tableElement.row.add({
                                DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                                __raw: result,
                                0: firstColumn,
                                1: referenceNumber,
                                2: uvpRequestId,
                                3: subrequesttypeFormatted,
                                4: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                                5: `<span>${vendorName}</span>`,
                                6: vendorCode,
                                7: vendorEmail,
                                8: `<span>${legalEntityLabel}</span>`,
                                9: assignedTo,
                                10: requestedBy,  // Requested By - column 10
                                11: assignedDateTimeHtml,  // Assigned Date - column 11
                                12: hoursRemainingHtml,  // Due Date - column 12
                                13: requestStatus,  // Approval Status - column 13
                                14: finalRequestStatusFormatted,  // Final Status - column 14
                                15: actionHtmlForPending,  // Action - column 15
                                16: prNumber,
                                17: poNumber,
                                18: ProjectID,
                                19: "",
                                20: EAuctionName,
                                21: "",
                                22: "",
                                23: "",
                                24: "",
                                25: ""
                            });
                            tableElement.column(0).visible(false);
                        } else {
                            // allReqTable and other tables don't have checkbox - use detected column positions
                            // If column indices were detected, use them; otherwise use default positions (8, 9, 10, 11)
                            let finalAssignedDateIdx = assignedDateIdx !== -1 ? assignedDateIdx : 8;
                            let finalDueDateIdx = dueDateIdx !== -1 ? dueDateIdx : 9;
                            let finalApprovalStatusIdx = approvalStatusIdx !== -1 ? approvalStatusIdx : 10;
                            let finalFinalStatusIdx = finalStatusIdx !== -1 ? finalStatusIdx : 11;

                            // Build action HTML: for Finance team (team type 25) show 3-dot menu with "View" option,
                            // for all other teams keep existing "Review" button + any PSSC reassign button
                            let actionHtml;
                            if (userWithTeam.al_teamtype == 25 && !isPendingTableForAll) {
                                let actions = [{
                                    label: "View",
                                    onclick: `window.location.href='/request-journey/?id=${requestId}&view=true'`
                                }];
                                let actionBtn = getActionMenuHTML(actions);
                                actionHtml = `${actionBtn}`;
                            } else {
                                // Use canReview to conditionally show Review button

                                if (canReview) {

                                    actionHtml = `<a href="/request-journey/?id=${requestId}" title="Review" class="btn btn-sm btn-secondary mr-1">Review</a>${psscReAssignBtn}`;

                                } else {

                                    // If can't review, show View button

                                    let actions = [{

                                        label: "View",

                                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`

                                    }];
                                    let actionBtn = getActionMenuHTML(actions);
                                    actionHtml = `${actionBtn}`;
                                }
                            }

                            let rowData = {
                                DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                                __raw: result,
                                0: referenceNumber,
                                1: uvpRequestId,
                                2: subrequesttypeFormatted,
                                3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                                4: `<span>${vendorName}</span>`,
                                5: vendorCode,
                                6: vendorEmail,
                                7: `<span>${legalEntityLabel}</span>`,
                                8: assignedTo,
                                9: requestedBy,
                                10: assignedDateTimeHtml,  // Assigned Date & Time (default position)
                                11: hoursRemainingHtml,    // Due Date (default position)
                                12: requestStatus,        // Approval Status (default position)
                                13: finalRequestStatusFormatted, // Final Status (default position)
                                14: `<div class="d-flex justify-content-center">${actionHtml}</div>`,
                                15: prNumber,
                                16: poNumber,
                                17: ProjectID,
                                18: "",
                                19: EAuctionName,
                                20: requestedBy,
                                21: assignedDateTimeHtml,
                                22: hoursRemainingHtml,
                                23: "",
                                24: ""
                            };

                            // If column positions were detected and differ from defaults, update them
                            if (finalAssignedDateIdx !== 8) {
                                rowData[finalAssignedDateIdx] = assignedDateTimeHtml;
                                delete rowData[8];
                            }
                            if (finalDueDateIdx !== 9) {
                                rowData[finalDueDateIdx] = hoursRemainingHtml;
                                delete rowData[9];
                            }
                            if (finalApprovalStatusIdx !== 10) {
                                rowData[finalApprovalStatusIdx] = requestStatus;
                                delete rowData[10];
                            }
                            if (finalFinalStatusIdx !== 11) {
                                rowData[finalFinalStatusIdx] = finalRequestStatusFormatted;
                                delete rowData[11];
                            }

                            tableElement.row.add(rowData);
                        }
                    }
                    // For non-pending tables: Set "Requested By" (column 8 for allReqTable) visibility based on team type
                    // For pendingReqTable, column 8 is Legal Entity (always visible), so skip this logic
                    if (!isPendingTableForAll) {
                        if (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3) {
                            tableElement.column(9).visible(true); // Requested By for allReqTable
                        } else {
                            tableElement.column(9).visible(false); // Requested By for allReqTable
                        }
                    }
                    else if (isPendingTableForAll) {
                        // Show "Requested By" for Line Manager (2) and PSSC (3) teams in pending table
                        const showRequestedBy = userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3);
                        tableElement.column(10).visible(showRequestedBy); // Requested By - VISIBLE FOR LINE MANAGER (2) AND PSSC (3) TEAMS
                    }
                }
            } catch (error) {
                console.log(error);
            }
        }

        // When appending to pending table, save current page before draw
        let currentPageToRestore = -1;
        if (append && isPendingTableForAll) {
            try {
                currentPageToRestore = tableElement.page.info().page;
            } catch (e) {
                // Ignore if page info not available
            }

            // Hide "no records found" message if it exists (to prevent it from showing during load)
            let $table = $(tableElement.table().node());
            let $emptyRow = $table.find('tbody tr.dataTables_empty');
            if ($emptyRow.length) {
                $emptyRow.hide();
            }
        }

        // Draw without resetting page
        tableElement.draw(false);

        // Restore page after draw if we were appending to pending table
        if (append && isPendingTableForAll && currentPageToRestore >= 0) {
            // Use requestAnimationFrame to ensure DOM is updated
            requestAnimationFrame(() => {
                setTimeout(() => {
                    try {
                        tableElement.page(currentPageToRestore).draw('page');
                    } catch (e) {
                        // Ignore errors
                    }
                }, 50);
            });
        }

        // Reorder columns: Move "Assigned Date" and "Due Date" to appear before "Approval Status" and "Final Status"
        // Use the variable already declared at the start of the function

        // For tables without checkbox (allReqTable, etc.), reorder columns using DOM manipulation
        if (!isPendingTableForAll) {
            try {
                let $table = $(tableElement.table().node());
                let $thead = $table.find('thead tr');
                let $tbody = $table.find('tbody');

                // Get all header cells
                let $headers = $thead.find('th');

                // Find indices of columns by their text content
                let assignedDateIdx = -1, dueDateIdx = -1, approvalStatusIdx = -1, finalStatusIdx = -1;

                $headers.each(function (index) {
                    let headerText = $(this).text().trim().toLowerCase();
                    if (headerText.includes('assigned date') || headerText.includes('assined date')) {
                        assignedDateIdx = index;
                    } else if (headerText.includes('due date')) {
                        dueDateIdx = index;
                    } else if (headerText.includes('approval status')) {
                        approvalStatusIdx = index;
                    } else if (headerText.includes('final status')) {
                        finalStatusIdx = index;
                    }
                });

                // If "Assigned Date" and "Due Date" are after "Approval Status" and "Final Status", reorder them
                if (assignedDateIdx > approvalStatusIdx && dueDateIdx > finalStatusIdx) {
                    // Move "Assigned Date" and "Due Date" to positions 8 and 9 (after "Requested By")
                    // Move "Approval Status" and "Final Status" to positions 10 and 11

                    // Reorder headers
                    let $assignedDateHeader = $headers.eq(assignedDateIdx).detach();
                    let $dueDateHeader = $headers.eq(dueDateIdx).detach();
                    let $approvalStatusHeader = $headers.eq(approvalStatusIdx).detach();
                    let $finalStatusHeader = $headers.eq(finalStatusIdx).detach();

                    // Insert in correct order: Assigned Date (8), Due Date (9), Approval Status (10), Final Status (11)
                    $headers.eq(7).after($assignedDateHeader); // After "Requested By" (index 7)
                    $assignedDateHeader.after($dueDateHeader);
                    $dueDateHeader.after($approvalStatusHeader);
                    $approvalStatusHeader.after($finalStatusHeader);

                    // Reorder data cells in all rows
                    $tbody.find('tr').each(function () {
                        let $row = $(this);
                        let $cells = $row.find('td');

                        let $assignedDateCell = $cells.eq(assignedDateIdx).detach();
                        let $dueDateCell = $cells.eq(dueDateIdx).detach();
                        let $approvalStatusCell = $cells.eq(approvalStatusIdx).detach();
                        let $finalStatusCell = $cells.eq(finalStatusIdx).detach();

                        // Insert in correct order
                        $cells.eq(7).after($assignedDateCell); // After "Requested By" (index 7)
                        $assignedDateCell.after($dueDateCell);
                        $dueDateCell.after($approvalStatusCell);
                        $approvalStatusCell.after($finalStatusCell);
                    });

                    // Redraw table
                    tableElement.draw();
                }
            } catch (e) {
                console.log("Column reordering failed:", e);
            }
        }

        // Show same columns as pending tab in all tabs
        // Ensure proper column order: Vendor Name -> Vendor Code -> Vendor Email -> Legal Entity
        // Hide: Requested By, Assigned Date, Due Date (same as pending tab)
        if (isPendingTableForAll) {
            // For pendingReqTable: Ensure correct column visibility and order
            // Column indices: 0=checkbox(hidden), 1=RequestID, 2=UVPRequestID, 3=RequestSubType, 4=Date&Time, 5=VendorName, 6=VendorCode, 7=VendorEmail, 8=LegalEntity, 9=AssignedTo, 10=RequestedBy, 11=AssignedDate, 12=DueDate, 13=ApprovalStatus, 14=FinalStatus, 15=Action
            tableElement.column(5).visible(true); // Vendor Name
            tableElement.column(6).visible(true); // Vendor Code
            tableElement.column(7).visible(true); // Vendor Email
            tableElement.column(8).visible(true); // Legal Entity
            tableElement.column(9).visible(true); // Assigned To
            // Show "Requested By" for Line Manager (2) and PSSC (3) teams in pending table
            const showRequestedBy = userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3);
            tableElement.column(10).visible(showRequestedBy); // Requested By - VISIBLE FOR LINE MANAGER (2) AND PSSC (3) TEAMS
            // For PSSC teams (teamtype 3), show Assigned Date and Due Date
            if (userWithTeam && userWithTeam.al_teamtype == 3) {
                tableElement.column(11).visible(true); // Assigned Date - VISIBLE FOR PSSC TEAMS
                tableElement.column(12).visible(true); // Due Date - VISIBLE FOR PSSC TEAMS
            } else {
                tableElement.column(11).visible(false); // Hide Assigned Date for non-PSSC teams
                tableElement.column(12).visible(false); // Hide Due Date for non-PSSC teams
            }
            tableElement.column(13).visible(true); // Approval Status - ALWAYS VISIBLE
            tableElement.column(14).visible(true); // Final Status - ALWAYS VISIBLE
            tableElement.column(15).visible(true); // Action - ALWAYS VISIBLE
        } else {
            // For other tables: Ensure correct column visibility and order
            // Column indices: 0=RequestID, 1=UVPRequestID, 2=RequestSubType, 3=Date&Time, 4=VendorName, 5=VendorCode, 6=VendorEmail, 7=LegalEntity, 8=AssignedTo
            tableElement.column(4).visible(true); // Vendor Name
            tableElement.column(5).visible(true); // Vendor Code
            tableElement.column(6).visible(true); // Vendor Email
            tableElement.column(7).visible(true); // Legal Entity
            tableElement.column(8).visible(true); // Assigned To
            // Show "Requested By" only for Line Manager (2) and PSSC (3) teams
            if (userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3)) {
                tableElement.column(9).visible(true); // Requested By - VISIBLE FOR LINE MANAGER AND PSSC TEAMS
            } else {
                tableElement.column(9).visible(false); // Hide Requested By for other teams
            }
            // Show Assigned Date and Due Date only for PSSC team (3) in non-pending tables
            if (userWithTeam && userWithTeam.al_teamtype == 3) {
                tableElement.column(10).visible(true); // Assigned Date - VISIBLE FOR PSSC TEAMS
                tableElement.column(11).visible(true); // Due Date - VISIBLE FOR PSSC TEAMS
            } else {
                tableElement.column(10).visible(false); // Hide Assigned Date for non-PSSC teams
                tableElement.column(11).visible(false); // Hide Due Date for non-PSSC teams
            }
        }

        // FINAL FIX: Ensure Vendor Code and Vendor Email are ALWAYS visible at the end
        // Also ensure Approval Status, Final Status, and Action are ALWAYS visible for pending table
        // Also ensure Assigned To is ALWAYS visible in ALL tabs
        // For PSSC teams, also show Assigned Date and Due Date
        // This overrides any conditional logic that might have hidden them
        if (isPendingTableForAll) {
            tableElement.column(6).visible(true); // Vendor Code - ALWAYS VISIBLE
            tableElement.column(7).visible(true); // Vendor Email - ALWAYS VISIBLE
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
            // Show "Requested By" for Line Manager (2) and PSSC (3) teams in pending table
            const showRequestedBy = userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3);
            tableElement.column(10).visible(showRequestedBy); // Requested By - VISIBLE FOR LINE MANAGER (2) AND PSSC (3) TEAMS
            // For PSSC teams (teamtype 3), show Assigned Date and Due Date
            if (userWithTeam && userWithTeam.al_teamtype == 3) {
                tableElement.column(11).visible(true); // Assigned Date - VISIBLE FOR PSSC TEAMS
                tableElement.column(12).visible(true); // Due Date - VISIBLE FOR PSSC TEAMS
            } else {
                tableElement.column(11).visible(false); // Hide Assigned Date for non-PSSC teams
                tableElement.column(12).visible(false); // Hide Due Date for non-PSSC teams
            }
            tableElement.column(13).visible(true); // Approval Status - ALWAYS VISIBLE
            tableElement.column(14).visible(true); // Final Status - ALWAYS VISIBLE
            tableElement.column(15).visible(true); // Action - ALWAYS VISIBLE
        } else {
            tableElement.column(5).visible(true); // Vendor Code - ALWAYS VISIBLE
            tableElement.column(6).visible(true); // Vendor Email - ALWAYS VISIBLE
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }

        tableElement.draw(); // Redraw to apply visibility changes

        // After drawing, ensure Requested By, Assigned Date, and Due Date headers are hidden for pending table
        // Also ensure Approval Status, Final Status, and Action columns are visible
        if (isPendingTableForAll) {
            try {
                // Explicitly show Assigned To, Approval Status, Final Status, and Action columns
                tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
                // Show "Requested By" only for Line Manager (2) and PSSC (3) teams in pending table
                if (userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3)) {
                    tableElement.column(10).visible(true); // Requested By - VISIBLE FOR LINE MANAGER AND PSSC TEAMS
                } else {
                    tableElement.column(10).visible(false); // Hide Requested By for other teams
                }
                // For PSSC teams (teamtype 3), show Assigned Date and Due Date
                if (userWithTeam && userWithTeam.al_teamtype == 3) {
                    tableElement.column(11).visible(true); // Assigned Date - VISIBLE FOR PSSC TEAMS
                    tableElement.column(12).visible(true); // Due Date - VISIBLE FOR PSSC TEAMS
                } else {
                    tableElement.column(11).visible(false); // Hide Assigned Date for non-PSSC teams
                    tableElement.column(12).visible(false); // Hide Due Date for non-PSSC teams
                }
                tableElement.column(13).visible(true); // Approval Status - ALWAYS VISIBLE
                tableElement.column(14).visible(true); // Final Status - ALWAYS VISIBLE
                tableElement.column(15).visible(true); // Action - ALWAYS VISIBLE

                let $table = $(tableElement.table().node());
                let $headers = $table.find('thead tr th');
                $headers.each(function (index) {
                    let headerText = $(this).text().trim().toLowerCase();
                    // Show/hide Requested By header based on team type
                    if (headerText.includes('requested by')) {
                        // Show for Line Manager (2) and PSSC (3) teams, hide for others
                        if (userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3)) {
                            $(this).show(); // Show the header for Line Manager and PSSC teams
                        } else {
                            $(this).hide(); // Hide the header for other teams
                            tableElement.column(index).visible(false); // Hide the column
                        }
                    } else if (headerText.includes('assigned date') || headerText.includes('assined date') || headerText.includes('due date')) {
                        // For PSSC teams, show these headers; for others, hide them
                        if (userWithTeam && userWithTeam.al_teamtype == 3) {
                            $(this).show(); // Show the header for PSSC teams
                        } else {
                            $(this).hide(); // Hide the header for non-PSSC teams
                            tableElement.column(index).visible(false); // Hide the column
                        }
                    }
                });
            } catch (e) {
                console.log("Error hiding extra headers:", e);
            }
        }
    }
    else if (teamType == 1) {    //For Requestor (BU)
        tableElement.clear();
        for (const result of reqList) {
            try {
                // Initialize column data variables to prevent undefined errors
                let assignedDateTimeHtml = "-";
                let hoursRemainingHtml = "-";
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                // let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";

                let reqStatus = result.al_requeststatus;
                let subReqType = result.al_subrequesttype;
                let requestId = result["al_requestid"];  // Guid
                //  let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;

                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;
                //added draft
                // Show Review for: New (1), Returned (5) ONLY when primary contact is logged-in user,
                // Rejected to BU (16), Draft (41), LM/GRC/etc. pending (2,18), and special 15 condition
                const canReview =
                    reqStatus == 1 ||
                    (reqStatus == 5 && result._al_primarycontact_value == loggedInUserId) ||
                    reqStatus == 16 ||
                    reqStatus == 41 ||
                    reqStatus == 2 ||
                    reqStatus == 18 ||
                    (reqStatus == 15 && subReqType != 8 && subReqType == 4);

                if (canReview) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "Review",
                        onclick: `window.location.href='/request-journey/?id=${requestId}'`
                    });
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted, 12: `<div class="d-flex justify-content-center"><a href="/register-new-request/?id=` + requestId + `&view=true" title="View" class="btn btn-sm btn-primary">view</a></div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    // $(rowAdded.node()).attr('id', requestId);
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                        __raw: result,
                        0: referenceNumber,
                        1: uvpRequestId,
                        2: subrequesttypeFormatted,
                        3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                        4: `<span>${vendorName}</span>`,
                        5: vendorCode,
                        6: vendorEmail,
                        7: `<span>${legalEntityLabel}</span>`,
                        8: assignedTo,
                        9: requestedBy,
                        10: assignedDateTimeHtml,
                        11: hoursRemainingHtml,
                        12: requestStatus,
                        13: finalRequestStatusFormatted,
                        14: `<div class="d-flex justify-content-center">${actionBtn}</div>`,
                        15: prNumber,
                        16: poNumber,
                        17: ProjectID,
                        18: "",
                        19: EAuctionName,
                        20: "",
                        21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted, 12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                    // Show "Requested By" only for Line Manager (2) and PSSC (3) teams
                    if (userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3)) {
                        tableElement.column(9).visible(true); // Requested By - VISIBLE FOR LINE MANAGER AND PSSC TEAMS
                    } else {
                        tableElement.column(9).visible(false); // Hide Requested By for other teams
                    }
                    // Show Assigned Date and Due Date only for PSSC team (3) in non-pending tables
                    if (userWithTeam && userWithTeam.al_teamtype == 3) {
                        tableElement.column(10).visible(true); // Assigned Date - VISIBLE FOR PSSC TEAMS
                        tableElement.column(11).visible(true); // Due Date - VISIBLE FOR PSSC TEAMS
                    } else {
                        tableElement.column(10).visible(false); // Hide Assigned Date for non-PSSC teams
                        tableElement.column(11).visible(false); // Hide Due Date for non-PSSC teams
                    }
                }

            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Also show Assigned Date and Due Date for PSSC teams (same as pending tab)
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
            // For PSSC teams (teamtype 3), show Assigned Date and Due Date
            if (userWithTeam && userWithTeam.al_teamtype == 3) {
                tableElement.column(11).visible(true); // Assigned Date - VISIBLE FOR PSSC TEAMS
                tableElement.column(12).visible(true); // Due Date - VISIBLE FOR PSSC TEAMS
            } else {
                tableElement.column(11).visible(false); // Hide Assigned Date for non-PSSC teams
                tableElement.column(12).visible(false); // Hide Due Date for non-PSSC teams
            }
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
            // Show Assigned Date and Due Date only for PSSC team (3) in non-pending tables
            if (userWithTeam && userWithTeam.al_teamtype == 3) {
                tableElement.column(10).visible(true); // Assigned Date - VISIBLE FOR PSSC TEAMS
                tableElement.column(11).visible(true); // Due Date - VISIBLE FOR PSSC TEAMS
            } else {
                tableElement.column(10).visible(false); // Hide Assigned Date for non-PSSC teams
                tableElement.column(11).visible(false); // Hide Due Date for non-PSSC teams
            }
        }
    }
    else if (teamType == 21) {    //For BU Head
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";

                let reqStatus = result.al_requeststatus;
                let subReqType = result.al_subrequesttype;
                let requestId = result["al_requestid"];  // Guid
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;

                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                let assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                let hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                //added draft
                // Show Review for: New (1), Returned (5) ONLY when primary contact is logged-in user,
                // Rejected to BU (16), Draft (41), LM/GRC/etc. pending (2,18), and special 15 condition
                const canReview =
                    reqStatus == 1 ||
                    (reqStatus == 5 && result._al_primarycontact_value == loggedInUserId) ||
                    reqStatus == 16 ||
                    reqStatus == 41 ||
                    reqStatus == 2 ||
                    reqStatus == 18 ||
                    (reqStatus == 15 && subReqType != 8 && subReqType == 4);

                if (canReview) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "Review",
                        onclick: `window.location.href='/request-journey/?id=${requestId}'`
                    });
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                        __raw: result,
                        0: referenceNumber,
                        1: uvpRequestId,
                        2: subrequesttypeFormatted,
                        3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                        4: `<span>${vendorName}</span>`,
                        5: vendorCode,
                        6: vendorEmail,
                        7: `<span>${legalEntityLabel}</span>`,
                        8: assignedTo,
                        9: requestedBy,
                        10: assignedDateTimeHtml,
                        11: hoursRemainingHtml,
                        12: requestStatus,
                        13: finalRequestStatusFormatted,
                        14: `<div class="d-flex justify-content-center">${actionBtn}</div>`,
                        15: prNumber,
                        16: poNumber,
                        17: ProjectID,
                        18: "",
                        19: EAuctionName,
                        20: "",
                        21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                        __raw: result,
                        0: referenceNumber,
                        1: uvpRequestId,
                        2: subrequesttypeFormatted,
                        3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                        4: `<span>${vendorName}</span>`,
                        5: vendorCode,
                        6: vendorEmail,
                        7: `<span>${legalEntityLabel}</span>`,
                        8: assignedTo,
                        9: requestedBy,
                        10: assignedDateTimeHtml,
                        11: hoursRemainingHtml,
                        12: requestStatus,
                        13: finalRequestStatusFormatted,
                        14: `<div class="d-flex justify-content-center">${actionBtn}</div>`,
                        15: prNumber,
                        16: poNumber,
                        17: ProjectID,
                        18: "",
                        19: EAuctionName,
                        20: "",
                        21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // BU Head specific column visibility - can be customized differently from BU
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }

            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 22) {    //For Finance Head
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";

                let reqStatus = result.al_requeststatus;
                let subReqType = result.al_subrequesttype;
                let requestId = result["al_requestid"];  // Guid
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;

                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                let assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                let hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                //added draft
                // Show Review for: New (1), Returned (5) ONLY when primary contact is logged-in user,
                // Rejected to BU (16), Draft (41), LM/GRC/etc. pending (2,18), and special 15 condition
                const canReview =
                    reqStatus == 1 ||
                    (reqStatus == 5 && result._al_primarycontact_value == loggedInUserId) ||
                    reqStatus == 16 ||
                    reqStatus == 41 ||
                    reqStatus == 2 ||
                    reqStatus == 18 ||
                    (reqStatus == 15 && subReqType != 8 && subReqType == 4);

                if (canReview) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "Review",
                        onclick: `window.location.href='/request-journey/?id=${requestId}'`
                    });
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                        __raw: result,
                        0: referenceNumber,
                        1: uvpRequestId,
                        2: subrequesttypeFormatted,
                        3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                        4: `<span>${vendorName}</span>`,
                        5: vendorCode,
                        6: vendorEmail,
                        7: `<span>${legalEntityLabel}</span>`,
                        8: assignedTo,
                        9: requestedBy,
                        10: assignedDateTimeHtml,
                        11: hoursRemainingHtml,
                        12: requestStatus,
                        13: finalRequestStatusFormatted,
                        14: `<div class="d-flex justify-content-center">${actionBtn}</div>`,
                        15: prNumber,
                        16: poNumber,
                        17: ProjectID,
                        18: "",
                        19: EAuctionName,
                        20: "",
                        21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                        __raw: result,
                        0: referenceNumber,
                        1: uvpRequestId,
                        2: subrequesttypeFormatted,
                        3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                        4: `<span>${vendorName}</span>`,
                        5: vendorCode,
                        6: vendorEmail,
                        7: `<span>${legalEntityLabel}</span>`,
                        8: assignedTo,
                        9: requestedBy,
                        10: assignedDateTimeHtml,
                        11: hoursRemainingHtml,
                        12: requestStatus,
                        13: finalRequestStatusFormatted,
                        14: `<div class="d-flex justify-content-center">${actionBtn}</div>`,
                        15: prNumber,
                        16: poNumber,
                        17: ProjectID,
                        18: "",
                        19: EAuctionName,
                        20: "",
                        21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Finance Head specific column visibility - can be customized differently from BU
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }

            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 24) {    //For SVP Procurement
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                //let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let subReqType = result.al_subrequesttype;
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"];  // Guid
                //let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";
                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequest?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                let assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                let hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");
                if (reqStatus == 7) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Vendor Code and Vendor Email should always be visible for all roles
                // For non-pending tables: Vendor Code (5), Vendor Email (6) - always visible
                tableElement.column(5).visible(true); // Vendor Code - ALWAYS VISIBLE
                tableElement.column(6).visible(true); // Vendor Email - ALWAYS VISIBLE
            } catch (error) {
                console.log(error);
            }
        }
        // FINAL FIX: Ensure Vendor Code and Vendor Email are ALWAYS visible
        tableElement.column(5).visible(true); // Vendor Code - ALWAYS VISIBLE
        tableElement.column(6).visible(true); // Vendor Email - ALWAYS VISIBLE
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 25) {    //For Finance
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";

                let reqStatus = result.al_requeststatus;
                let subReqType = result.al_subrequesttype;
                let requestId = result["al_requestid"];  // Guid
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;

                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");
                //added draft
                // Show Review for: New (1), Returned (5) ONLY when primary contact is logged-in user,
                // Rejected to BU (16), Draft (41), LM/GRC/etc. pending (2,18), and special 15 condition
                const canReview =
                    reqStatus == 1 ||
                    (reqStatus == 5 && result._al_primarycontact_value == loggedInUserId) ||
                    reqStatus == 16 ||
                    reqStatus == 41 ||
                    reqStatus == 2 ||
                    reqStatus == 18 ||
                    (reqStatus == 15 && subReqType != 8 && subReqType == 4);

                if (canReview) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "Review",
                        onclick: `window.location.href='/request-journey/?id=${requestId}'`
                    });
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Finance specific column visibility - can be customized differently from BU
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }

            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 23) {    //For Category Manager
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";

                let reqStatus = result.al_requeststatus;
                let subReqType = result.al_subrequesttype;
                let requestId = result["al_requestid"];  // Guid
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;

                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                let assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                let hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                //added draft
                // Show Review for: New (1), Returned (5) ONLY when primary contact is logged-in user,
                // Rejected to BU (16), Draft (41), LM/GRC/etc. pending (2,18), and special 15 condition
                const canReview =
                    reqStatus == 1 ||
                    (reqStatus == 5 && result._al_primarycontact_value == loggedInUserId) ||
                    reqStatus == 16 ||
                    reqStatus == 41 ||
                    reqStatus == 2 ||
                    reqStatus == 18 ||
                    (reqStatus == 15 && subReqType != 8 && subReqType == 4);

                if (canReview) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "Review",
                        onclick: `window.location.href='/request-journey/?id=${requestId}'`
                    });
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Category Manager specific column visibility - can be customized differently from BU
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }

            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 2) {   //For Line Manager
        tableElement.clear();
        for (const result of reqList) {
            try {
                // Initialize column data variables to prevent undefined errors
                let assignedDateTimeHtml = "-";
                let hoursRemainingHtml = "-";
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                //let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let subReqType = result.al_subrequesttype;
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"]; // Guid
                //let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A"
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }

                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                if (reqStatus == 2) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    if (result['_al_primarycontact_value'] == userWithTeam.contactid) {
                        actions.push({
                            label: "Approve",
                            onclick: `approveRequest('${teamType}','${requestId}')`
                        });
                    }

                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: {
                    //         'title': reassignDetail, 'data-auctionname': result.al_eauctionname,
                    //         'data-projectid': result.al_projectnameid
                    //     }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(0).visible(true);
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                    tableElement.column(9).visible(true); // Assigned To
                    tableElement.column(10).visible(true); // Assigned To
                    tableElement.column(11).visible(true); // Assigned To
                }
            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Reorder columns: Move "Assigned Date" and "Due Date" to appear before "Approval Status" and "Final Status"
        // Check which table we're working with
        let tableIdForResolved = tableElement.table().node().id;
        let isPendingTableForResolved = tableIdForResolved === 'pendingReqTable';

        // For tables without checkbox (grcApproveTable, etc.), reorder columns using DOM manipulation
        if (!isPendingTableForResolved) {
            try {
                // First, ensure columns 8 and 9 are visible temporarily for reordering
                let wasCol8Visible = tableElement.column(8).visible();
                let wasCol9Visible = tableElement.column(9).visible();
                tableElement.column(8).visible(true);
                tableElement.column(9).visible(true);
                tableElement.draw();

                // Force a synchronous DOM update by accessing the table node
                let tableNode = tableElement.table().node();
                let $table = $(tableNode);
                let $thead = $table.find('thead tr');
                let $tbody = $table.find('tbody');

                // Get all header cells
                let $headers = $thead.find('th');

                // Find indices by header text - we know the expected structure
                // Expected: 0-7 (various), 8 (Assigned Date), 9 (Due Date), 10 (Approval Status), 11 (Final Status), 12 (Action)
                let assignedDateIdx = -1, dueDateIdx = -1, approvalStatusIdx = -1, finalStatusIdx = -1, requestedByIdx = -1;

                $headers.each(function (index) {
                    let headerText = $(this).text().trim().toLowerCase();
                    if (headerText.includes('requested by')) {
                        requestedByIdx = index;
                    } else if (headerText.includes('assigned date') || headerText.includes('assined date')) {
                        assignedDateIdx = index;
                    } else if (headerText.includes('due date')) {
                        dueDateIdx = index;
                    } else if (headerText.includes('approval status')) {
                        approvalStatusIdx = index;
                    } else if (headerText.includes('final status')) {
                        finalStatusIdx = index;
                    }
                });

                // If we found the columns and they're not in the correct order, reorder them
                if (requestedByIdx !== -1) {
                    // If Assigned Date and Due Date columns don't exist in headers, the data is misaligned
                    // We need to reorder based on where the data actually is (indices 8, 9, 10, 11)
                    if (assignedDateIdx === -1 && dueDateIdx === -1 && $headers.length >= 12) {
                        // Headers don't have Assigned Date/Due Date, but data does
                        // Reorder: take columns 8, 9, 10, 11 and ensure they're in the right order
                        let $col8Header = $headers.eq(8).detach();
                        let $col9Header = $headers.eq(9).detach();
                        let $col10Header = $headers.eq(10).detach();
                        let $col11Header = $headers.eq(11).detach();

                        // Get the "Requested By" header
                        let $requestedByHeader = $headers.eq(requestedByIdx);

                        // Insert in correct order: Assigned Date (8), Due Date (9), Approval Status (10), Final Status (11)
                        $requestedByHeader.after($col8Header);
                        $col8Header.after($col9Header);
                        $col9Header.after($col10Header);
                        $col10Header.after($col11Header);

                        // Reorder data cells in all rows
                        $tbody.find('tr').each(function () {
                            let $row = $(this);
                            let $cells = $row.find('td');

                            if ($cells.length >= 12) {
                                let $col8Cell = $cells.eq(8).detach();
                                let $col9Cell = $cells.eq(9).detach();
                                let $col10Cell = $cells.eq(10).detach();
                                let $col11Cell = $cells.eq(11).detach();

                                let $requestedByCell = $cells.eq(requestedByIdx);

                                $requestedByCell.after($col8Cell);
                                $col8Cell.after($col9Cell);
                                $col9Cell.after($col10Cell);
                                $col10Cell.after($col11Cell);
                            }
                        });
                    } else if (assignedDateIdx !== -1 && dueDateIdx !== -1 && approvalStatusIdx !== -1 && finalStatusIdx !== -1) {
                        // Headers exist but might be in wrong order
                        if (assignedDateIdx !== requestedByIdx + 1 || dueDateIdx !== requestedByIdx + 2) {
                            // Detach and reorder
                            let indicesToReorder = [assignedDateIdx, dueDateIdx, approvalStatusIdx, finalStatusIdx].sort((a, b) => b - a);
                            let $detachedHeaders = indicesToReorder.map(idx => $headers.eq(idx).detach());

                            let $requestedByHeader = $headers.eq(requestedByIdx);
                            let correctOrder = [assignedDateIdx, dueDateIdx, approvalStatusIdx, finalStatusIdx];

                            correctOrder.forEach(function (originalIdx) {
                                let headerIndex = indicesToReorder.indexOf(originalIdx);
                                $requestedByHeader.after($detachedHeaders[headerIndex]);
                                $requestedByHeader = $detachedHeaders[headerIndex];
                            });

                            // Reorder data cells
                            $tbody.find('tr').each(function () {
                                let $row = $(this);
                                let $cells = $row.find('td');
                                let $detachedCells = indicesToReorder.map(idx => $cells.eq(idx).detach());
                                let $requestedByCell = $cells.eq(requestedByIdx);

                                correctOrder.forEach(function (originalIdx) {
                                    let cellIndex = indicesToReorder.indexOf(originalIdx);
                                    $requestedByCell.after($detachedCells[cellIndex]);
                                    $requestedByCell = $detachedCells[cellIndex];
                                });
                            });
                        }
                    }

                    // Redraw table
                    tableElement.draw();
                }

                // Restore original visibility settings
                tableElement.column(8).visible(wasCol8Visible);
                tableElement.column(9).visible(wasCol9Visible);
                tableElement.draw();
            } catch (e) {
                console.log("Column reordering failed for Resolved tab:", e);
            }
        }

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 3) {  //For PSSC
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                //let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let subReqType = result.al_subrequesttype;
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"];  // Guid
                // let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }

                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }

                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                let tooltipText = `PO Number: ${poNumber || 'N/A'}\nPR Number: ${prNumber || 'N/A'}\nVendor Name: ${vendorName || 'N/A'}`;
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                if (reqStatus == 3 && result._al_primarycontact_value == $("#userIdHidden").html()) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: `<span title="` + prNumber + `">` + prNumber + `</span>`, 16: `<span title="` + poNumber + `">` + poNumber + `</span>`, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: `<span title="` + prNumber + `">` + prNumber + `</span>`, 7: `<span title="` + poNumber + `">` + poNumber + `</span>`, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: `<span title="` + prNumber + `">` + prNumber + `</span>`, 16: `<span title="` + poNumber + `">` + poNumber + `</span>`, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowA dded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: `<span title="` + prNumber + `">` + prNumber + `</span>`, 7: `<span title="` + poNumber + `">` + poNumber + `</span>`, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }
            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 4) {  //For GRC
        tableElement.clear();
        for (const result of reqList) {
            try {
                // Initialize column data variables to prevent undefined errors
                let assignedDateTimeHtml = "-";
                let hoursRemainingHtml = "-";
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                // let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let subReqType = result.al_subrequesttype;
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"];  // Guid
                // let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }

                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequest?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }

                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                if (reqStatus == 7) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }
            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 5) {  //For SVP
        tableElement.clear();
        for (const result of reqList) {
            try {
                // Initialize column data variables to prevent undefined errors
                let assignedDateTimeHtml = "-";
                let hoursRemainingHtml = "-";
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                //let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let subReqType = result.al_subrequesttype;
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"];  // Guid
                //let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";
                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequest?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                if (reqStatus == 15) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actions.push({
                        label: "Approve",
                        onclick: `approveRequest('${teamType}','${requestId}')`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }
            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 6) {  //For CEEPO
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                // let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let subReqType = result.al_subrequesttype;
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"];  // Guid
                // let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A"
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequest?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                let assignedDateTime = getAssignedDateTime(result);
                let dueDate = await calculateDueDate(result);
                let assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                let hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;

                if (reqStatus == 17) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actions.push({
                        label: "Approve",
                        onclick: `approveRequest('${teamType}','${requestId}')`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });
                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }
            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 7) {  //For PSSC Manager
        $('#profile-tab').text('Reassignment Requests');
        $('#home-tab').text('All');
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                // let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"]; // Guid
                let subReqType = result.al_subrequesttype;

                // let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";
                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span >`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span >`;
                if (reqStatus == 3) {
                    let rowAdded;
                    let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || "");
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "Re Assign ",
                        onclick: `openAssignModal(
                        '${result["_al_primarycontact_value"]}',  
                        '${requestId}','${result['_al_currentapprovalrequest_value']}',
                        '${result["al_referencenumber"]}',
                        '${result["al_vendorname"]}',
                        '${result["al_legalentitydetails"]}',
                        '${assignedTo}',
                        '${result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["createdon@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_priority@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_changetype@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_requeststatus@OData.Community.Display.V1.FormattedValue"]}'
                        )`
                    });
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let firstColumn = "";
                    // tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: uvpRequestId, 3: subrequesttypeFormatted, 4: createdOn, 5: `<span>${vendorName}</span>`, 6: `<span>${legalEntityLabel}</span>`, 7: assignedTo, 8: requestedBy, 9: requestStatus, 10: finalRequestStatusFormatted, 11: `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`, 12: prNumber, 13: poNumber,
                    //     14: ProjectID, 15: EAuctionName
                    // });
                    //tableElement.column(0).visible(false);
                    // Vendor Code and Vendor Email should always be visible for all roles
                    // For non-pending tables: Vendor Code (5), Vendor Email (6) - always visible
                    tableElement.column(5).visible(true); // Vendor Code - ALWAYS VISIBLE
                    tableElement.column(6).visible(true); // Vendor Email - ALWAYS VISIBLE
                }
            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 8) {   //For FSSC Line Manager
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                //let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let subReqType = result.al_subrequesttype;
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"];  // Guid
                // let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A";
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";
                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }
                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";

                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || result["al_reassigndetailfssc"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || result["al_reassigndetailfssc"] || "");

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                // Initialize with default values to prevent undefined errors
                let assignedDateTimeHtml = "-";
                let hoursRemainingHtml = "-";
                try {
                    let assignedDateTime = getAssignedDateTime(result);
                    let dueDate = await calculateDueDate(result);
                    assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                    hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;
                } catch (calcError) {
                    console.warn("Error calculating assigned date/time or due date:", calcError);
                    // Variables already initialized with "-" above
                }

                if (reqStatus == 18) {  //Assign FSSC Team Member
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "Re Assign ",
                        onclick: `openAssignModalForFSSC(
                        '${result["_al_primarycontact_value"]}',
                        '${requestId}',
                        '${result['_al_currentapprovalrequest_value']}',
                        '${result["al_referencenumber"]}',
                        '${result["al_vendorname"]}',
                        '${result["al_legalentitydetails"]}',
                        '${assignedTo}',
                        '${result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["createdon@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_priority@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_changetype@OData.Community.Display.V1.FormattedValue"]}',
                        '${result["al_requeststatus@OData.Community.Display.V1.FormattedValue"]}'
                        )`
                    });
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    // Vendor Code and Vendor Email should always be visible for all roles
                    // For non-pending tables: Vendor Code (5), Vendor Email (6) - always visible
                    tableElement.column(5).visible(true); // Vendor Code - ALWAYS VISIBLE
                    tableElement.column(6).visible(true); // Vendor Email - ALWAYS VISIBLE
                }
                else if (reqStatus == 19) {  //Assige FSSC Manager
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actions.push({
                        label: "Approve",
                        onclick: `approveRequest('${teamType}','${requestId}')`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                }
                else if (reqStatus != 42) {//Not in draft
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: "window.location.href='/request-journey/?id=" + requestId + "&view=true'"
                        // onclick: `window.location.href='/request-journey/?id=` + requestId + `'`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                }
                // Show all columns for all roles
                if (isPendingTableForAll) {
                    tableElement.column(5).visible(true); // Vendor Name
                    tableElement.column(6).visible(true); // Vendor Code
                    tableElement.column(7).visible(true); // Vendor Email
                    tableElement.column(8).visible(true); // Legal Entity
                    tableElement.column(9).visible(true); // Assigned To
                } else {
                    tableElement.column(4).visible(true); // Vendor Name
                    tableElement.column(5).visible(true); // Vendor Code
                    tableElement.column(6).visible(true); // Vendor Email
                    tableElement.column(7).visible(true); // Legal Entity
                    tableElement.column(8).visible(true); // Assigned To
                }
            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }
    else if (teamType == 9) { //For FSSC Team Member
        tableElement.clear();
        for (const result of reqList) {
            try {
                let EAuctionName = result.al_eauctionname ? result.al_eauctionname : "";
                let ProjectID = result.al_projectnameid ? result.al_projectnameid : "";
                // let assignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] ? result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] : "";
                let subReqType = result.al_subrequesttype;
                let reqStatus = result.al_requeststatus;
                let requestId = result["al_requestid"];  // Guid
                // let referenceNumber = result["al_referencenumber"]; // Text
                let redirectLink = subReqType == 13 || subReqType == 14 ? `/request-journey/?id=${requestId}&view=true` : `/request-journey/?id=${requestId}`;
                let uvpRequestId = '-';
                let uvpVendorRequestId = null;
                if (result.al_UVPVendorRequest) {
                    uvpRequestId = result.al_UVPVendorRequest.ag_requestnumber;
                    uvpVendorRequestId = result.al_UVPVendorRequest.ag_uvpvendorrequestid || null;
                } else if (result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest && result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest.length > 0) {
                    let uvpApproval = result.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest[0];
                    uvpRequestId = uvpApproval.ag_requestnumber;
                    uvpVendorRequestId = (uvpApproval && (uvpApproval.ag_uvpvendorrequestid || (uvpApproval.al_UVPVendorRequest && uvpApproval.al_UVPVendorRequest.ag_uvpvendorrequestid))) || null;
                }
                uvpRequestId = formatUvpRequestIdCell(uvpRequestId, uvpVendorRequestId);
                let referenceNumber = `<a class="table_head_color" href="${redirectLink}"><span class="table_head_color">${result["al_referencenumber"]}</span></a>`;
                let requesttypeFormatted = result["al_requesttype@OData.Community.Display.V1.FormattedValue"];
                let subrequesttypeFormatted = result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"];
                let vendorName = "N/A";
                let eAuctionName = "N/A";
                if (result["al_vendormultiplelabel"]) {
                    vendorName = result["al_vendormultiplelabel"];
                } else if (result["al_vendorname"]) {
                    vendorName = result["al_vendorname"];
                } else if (result["al_eauctionname"]) {
                    vendorName = "N/A"
                }
                if (result["al_eauctionname"]) {
                    eAuctionName = result["al_eauctionname"]
                }
                else {
                    eAuctionName = "N/A";
                }
                // Extract Vendor Code and Vendor Email
                let vendorCode = "N/A";
                let vendorEmail = "N/A";
                if (result.al_UVPVendorRequest) {
                    vendorCode = result.al_UVPVendorRequest.ag_vendorcode || "N/A";
                    vendorEmail = result.al_UVPVendorRequest.ag_vendorcontactemail || "N/A";
                }
                let createdOn = result["createdon@OData.Community.Display.V1.FormattedValue"] ? result["createdon@OData.Community.Display.V1.FormattedValue"] : "";//datetime in text
                let createdOnTimestamp = result["createdon"] ? new Date(result["createdon"]).getTime() : 0; // Raw timestamp for sorting
                let legalEntityLabel = result["al_legalentitydetails"] ? result["al_legalentitydetails"] : "N/A"; // Text
                let poNumber = result.al_ponumber ? result.al_ponumber : "";
                let prNumber = result.al_prnumber ? result.al_prnumber : "";
                let requestStatus = result["al_requeststatus@OData.Community.Display.V1.FormattedValue"];
                requestStatus = requestStatus != null ? requestStatus : "";

                let finalRequestStatusFormatted = result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"];
                finalRequestStatusFormatted = finalRequestStatusFormatted != null ? finalRequestStatusFormatted : "";
                let requestedBy = result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                let reqStatusClass = requestStatus.replace(/ /g, '');
                let finalReqStatusClass = finalRequestStatusFormatted.replace(/ /g, '');

                requestStatus = `<span class="approvalStatus ${reqStatusClass}">${requestStatus}</span>`;
                finalRequestStatusFormatted = `<span class="${finalReqStatusClass}">${finalRequestStatusFormatted}</span>`;
                let reassignDetail = eAuctionName != 'N/A' ? (result["al_reassigndetail"] || result["al_reassigndetailfssc"] || "") + " E-auction Name: " + eAuctionName : (result["al_reassigndetail"] || result["al_reassigndetailfssc"] || "");
                let rawAssignedTo = result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "";
                let assignedTo = rawAssignedTo;
                if ([13, 14].includes(subReqType) && !rawAssignedTo) {
                    let approvalRequests = result?.al_approvalrequest_Request_al_request;
                    let checkreq = approvalRequests?.filter(item => item._al_request_value == requestId && item.al_requeststatus == 1 && item._al_team_value != null);
                    if (checkreq?.length != 0) {
                        assignedTo = `<a href="javascript:void(0);" class="assignedToClick text-primary check-pending" data-id="${requestId}">UVP Team</a>`;
                    }
                    else {
                        assignedTo = `<span>-</span>`;
                    }
                }

                // Calculate assigned date/time and due date for new columns (PSSC team only)
                // Initialize with default values to prevent undefined errors
                let assignedDateTimeHtml = "-";
                let hoursRemainingHtml = "-";
                try {
                    let assignedDateTime = getAssignedDateTime(result);
                    let dueDate = await calculateDueDate(result);
                    assignedDateTimeHtml = assignedDateTime.timestamp > 0 ? `<span data-order="${assignedDateTime.timestamp}">${assignedDateTime.formatted}</span>` : assignedDateTime.formatted;
                    hoursRemainingHtml = dueDate.timestamp > 0 ? `<span data-order="${dueDate.timestamp}">${dueDate.formatted}</span>` : dueDate.formatted;
                } catch (calcError) {
                    console.warn("Error calculating assigned date/time or due date:", calcError);
                    // Variables already initialized with "-" above
                }

                if (reqStatus == 18) {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid },
                        __raw: result,
                        0: referenceNumber,
                        1: uvpRequestId,
                        2: subrequesttypeFormatted,
                        3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`,
                        4: `<span>${vendorName}</span>`,
                        5: vendorCode,
                        6: vendorEmail,
                        7: `<span>${legalEntityLabel}</span>`,
                        8: assignedTo,
                        9: requestedBy,
                        10: assignedDateTimeHtml,
                        11: hoursRemainingHtml,
                        12: requestStatus,
                        13: finalRequestStatusFormatted,
                        14: `<div class="d-flex justify-content-center">${actionBtn}</div>`,
                        15: prNumber,
                        16: poNumber,
                        17: ProjectID,
                        18: "",
                        19: EAuctionName,
                        20: "",
                        21: ""
                    });
                    // let rowAdded = tableElement.row.add({
                    //     DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted,
                    //     12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName
                    // });
                    $(rowAdded.node()).attr('id', requestId);
                }
                else {
                    let actions = [];
                    let actionBtn = "";
                    actions.push({
                        label: "View",
                        onclick: subReqType == 13 || subReqType == 14 ? `window.location.href='/request-journey/?id=${requestId}&view=true'` : `window.location.href='/request-journey/?id=${requestId}&view=true'`
                    });

                    actionBtn = getActionMenuHTML(actions);
                    let rowAdded = tableElement.row.add({
                        DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: `<span data-order="${createdOnTimestamp}">${createdOn}</span>`, 4: `<span>${vendorName}</span>`, 5: vendorCode, 6: vendorEmail, 7: `<span>${legalEntityLabel}</span>`, 8: assignedTo, 9: requestedBy, 10: assignedDateTimeHtml, 11: hoursRemainingHtml, 12: requestStatus, 13: finalRequestStatusFormatted, 14: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 15: prNumber, 16: poNumber, 17: ProjectID,
                        18: "", 19: EAuctionName, 20: "", 21: ""
                    });
                    // let rowAdded = tableElement.row.add({ DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: requesttypeFormatted, 3: subrequesttypeFormatted, 4: legalEntityLabel, 5: vendorName, 6: prNumber, 7: poNumber, 8: createdOn, 9: assignedTo, 10: requestStatus, 11: finalRequestStatusFormatted, 12: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 13: ProjectID, 14: EAuctionName });
                    $(rowAdded.node()).attr('id', requestId);
                }
            } catch (error) {
                console.log(error);
            }
        }
        tableElement.draw();

        // Ensure Assigned To is ALWAYS visible in all tabs
        // Column 8 is "Assigned To" for non-pending tables, Column 9 is "Assigned To" for pending table
        // Do NOT hide "Assigned To" - it should always be visible
        if (isPendingTableForAll) {
            tableElement.column(9).visible(true); // Assigned To - ALWAYS VISIBLE
        } else {
            tableElement.column(8).visible(true); // Assigned To - ALWAYS VISIBLE IN ALL TABS
        }
    }

    // FINAL FIX: Ensure Vendor Code and Vendor Email are ALWAYS visible at the end of BindRequest
    // Also ensure Assigned Date and Due Date are ALWAYS hidden in non-pending tables (All/Resolved/Returned/Rejected/P1 tabs)
    // Also ensure "Requested By" is only visible for Line Manager (2) and PSSC (3) teams in non-pending tables
    // This overrides any conditional logic that might have hidden them
    // Check if this is pending table or other table
    let isPendingTable = tableElement.table().node().id === 'pendingReqTable';
    if (isPendingTable) {
        tableElement.column(6).visible(true); // Vendor Code - ALWAYS VISIBLE
        tableElement.column(7).visible(true); // Vendor Email - ALWAYS VISIBLE
        // Show "Requested By" only for Line Manager (2) and PSSC (3) teams in pending table
        if (userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3)) {
            tableElement.column(10).visible(true); // Requested By - VISIBLE FOR LINE MANAGER AND PSSC TEAMS
        } else {
            tableElement.column(10).visible(false); // Hide Requested By for other teams
        }
    } else {
        tableElement.column(5).visible(true); // Vendor Code - ALWAYS VISIBLE
        tableElement.column(6).visible(true); // Vendor Email - ALWAYS VISIBLE
        // Show Assigned Date and Due Date only for PSSC team (3) in non-pending tables (All/Resolved/Returned/Rejected/P1 tabs)
        if (userWithTeam && userWithTeam.al_teamtype == 3) {
            tableElement.column(10).visible(true); // Assigned Date - VISIBLE FOR PSSC TEAMS
            tableElement.column(11).visible(true); // Due Date - VISIBLE FOR PSSC TEAMS
        } else {
            tableElement.column(10).visible(false); // Hide Assigned Date for non-PSSC teams
            tableElement.column(11).visible(false); // Hide Due Date for non-PSSC teams
        }
        // Show "Requested By" only for Line Manager (2) and PSSC (3) teams in non-pending tables
        if (userWithTeam && (userWithTeam.al_teamtype == 2 || userWithTeam.al_teamtype == 3)) {
            tableElement.column(9).visible(true); // Requested By - VISIBLE FOR LINE MANAGER AND PSSC TEAMS
        } else {
            tableElement.column(9).visible(false); // Hide Requested By for other teams
        }
    }
    tableElement.draw(); // Redraw to apply final visibility changes
}


$(document).on('click', '.assignedToClick', function () {
    const requestId = $(this).data('id') || "Unknown";
    // Clear and show modal immediately with loading text
    $('#statusModal .modal-body').html(`<span class="text-muted">Loading team members...</span>`);
    $('#statusModal').modal('show');

    webapi.safeAjax({
        type: "GET",
        url: "/_api/al_approvalrequests?$select=al_approvalrequestid,_al_request_value,_al_team_value,al_teamtype&$filter=(_al_request_value eq " + requestId + " and al_requeststatus eq 1)",
        // url: "/_api/al_approvalrequests?$select=createdon,_al_request_value,_al_team_value,al_teamtype&$filter=_al_request_value eq " + requestId,
        contentType: "application/json",
        headers: {
            "Prefer": "odata.include-annotations=*"
        },
        success: function (data) {
            if (!data.value || data.value.length === 0) {
                // $('#statusModal .modal-body').html(`<span class="text-green">This request is completed.</span>`);
                $(`a.assignedToClick[data-id="${requestId}"]`).closest('td').html('-');
                return;
            }

            // Get the latest team based on created date
            let latestTeam = data.value.reduce((a, b) =>
                new Date(a.createdon) > new Date(b.createdon) ? a : b
            );

            let teamId = latestTeam._al_team_value;

            if (!teamId) {
                $('#statusModal .modal-body').html(`<span class="text-danger">Team not assigned.</span>`);
                return;
            }

            // Get contacts from the team
            webapi.safeAjax({
                type: "GET",
                url: "/_api/contacts?$select=fullname&$filter=_ag_uvpteam_value eq " + teamId,
                contentType: "application/json",
                headers: {
                    "Prefer": "odata.include-annotations=*"
                },
                success: function (contactData) {
                    let members = contactData.value;

                    if (!members || members.length === 0) {
                        $('#statusModal .modal-body').html(`<span class="text-warning">No members found for this team.</span>`);
                        return;
                    }

                    // Build line-by-line list of names
                    let namesList = members.map(member => `<div>${member.fullname}</div>`).join("");
                    $('#statusModal .modal-body').html(namesList);
                },
                error: function (xhr) {
                    console.error(xhr);
                    $('#statusModal .modal-body').html(`<span class="text-danger">Failed to load team members.</span>`);
                }
            });
        },
        error: function (xhr) {
            console.error(xhr);
            $('#statusModal .modal-body').html(`<span class="text-danger">Failed to load request data.</span>`);
        }
    });
});


$("#closeModal").on("click", function () {
    $('#statusModal').modal('hide');
});

function getActionMenuHTML(actions) {
    let menuItems = actions.map(action => {
        return `<li class="nav-item"><a class="dropdown-item" href="#" onclick="${action.onclick}">${action.label}</a></li>`;
    }).join("");

    return `
        <div class="dropdown">
            <button class="p-0 border-0 bg-transparent" type="button" data-bs-toggle="dropdown" aria-expanded="false">
                <i class="fas fa-ellipsis-h fs-6"></i>
            </button>
            <ul class="dropdown-menu dropdown-menu-end">
                ${menuItems}
            </ul>
        </div>`;
}

//Updates action by actor in the request and call method to update request
//@params: loggedInUserTeamType (to update action by actor based on team type), requestRecordId (Record id of the request)
async function approveRequest(loggedInUserTeamType, requestRecordId) {
    let isRequestUpdated = false;
    $('.custom-loader').css("display", "flex");

    let requestRecord = {};
    if (loggedInUserTeamType == 5) {   //For SVP
        requestRecord.al_actionbyactor = 10 //Approve By Svp
    } else if (loggedInUserTeamType == 2) {  //For Line Manager
        requestRecord.al_actionbyactor = 2//Approve By Line Manager
    } else if (loggedInUserTeamType == 6) {  //For CEEPO
        requestRecord.al_actionbyactor = 13 //Approved By CEEPO
    } else if (loggedInUserTeamType == 8) { // For FSSC Line Manager
        requestRecord.al_actionbyactor = 17 //Approved By FSSC Line Manager
    }
    else if (loggedInUserTeamType == 9) {  //For FSSC Team Member
        requestRecord.al_actionbyactor = 16 //Approved By FSSC Team
    }
    isRequestUpdated = await updateEntityRecord(requestRecord, "al_requests", requestRecordId);

    if (isRequestUpdated) {
        console.log("Request Approved");
        $('.custom-loader').css("display", "none");

        window.location.pathname = 'my-request';
    }
}

var totalForCount = [];
var returnReqForCount = [];
var completedForCount = [];
var pendingForCount = [];
let currentCustomSection = "";

function getRequestCount(requests) {
    totalForCount = requests;
    // Calculate pending requests: status 2, 41, or 5 (when primary contact matches logged-in user)
    // Ensure loggedInUserId is set (convert to string for consistent comparison)
    const userIdForComparison = loggedInUserId ? String(loggedInUserId).trim() : null;
    pendingForCount = requests.filter(item => {

        const status = item.al_requeststatus;

        const primaryContact = item._al_primarycontact_value ? String(item._al_primarycontact_value).trim() : null;
        // Status 2 or 41: always pending
        if (status == 2 || status == 41) {
            return true;
        }
        // Status 5: only pending if primary contact matches logged-in user
        if (status == 5 && userIdForComparison && primaryContact === userIdForComparison) {
            return true;
        }
        return false;
    });
    returnReqForCount = requests.filter(item => item.al_requeststatus == 5);
    // Calculate completed/resolved requests: status 6 or 40
    completedForCount = requests.filter(item => item.al_requeststatus == 6 || item.al_requeststatus == 40);
    var totalReq = requests.length;
    var pendingReq = pendingForCount.length;
    var returnReq = returnReqForCount.length;
    var completedReq = completedForCount.length;
    $(".totalReq").text(totalReq);
    $(".pendingReq").text(pendingReq);
    $(".returnReq").text(returnReq);
    $(".completedReq").text(completedReq);

}

$("#totalReq").on("change", function () {
    $(this).parent().find(".customdate").remove();
    const val = $(this).val();
    if (parseInt(val) === 4) {
        currentCustomSection = "totalReq";
        $('#customDateModal').modal('show');
    }
    else if ($(this).val() == 0) {
        $(".totalReq").text(totalForCount.length || 0);
    } else {
        const days = getDaysFromDropdown(val);
        const filtered = filterByDays(totalForCount, days);
        $(".totalReq").text(filtered.length || 0);
    }
});

$("#pendingReq").on("change", function () {
    $(this).parent().find(".customdate").remove();
    const val = $(this).val();
    if (parseInt(val) === 4) {
        currentCustomSection = "pendingReq";
        $('#customDateModal').modal('show');
    }
    else if ($(this).val() == 0) {
        $(".pendingReq").text(pendingForCount.length || 0);
    } else {
        const days = getDaysFromDropdown(val);
        const filtered = filterByDays(pendingForCount, days);
        $(".pendingReq").text(filtered.length || 0);
    }
});

$("#completedReq").on("change", function () {
    $(this).parent().find(".customdate").remove();
    const val = $(this).val();
    if (parseInt(val) === 4) {
        currentCustomSection = "completedReq";
        $('#customDateModal').modal('show');
    }
    else if ($(this).val() == 0) {
        $(".completedReq").text(completedForCount.length || 0);
    } else {
        const days = getDaysFromDropdown(val);
        const filtered = filterByDays(completedForCount, days);
        $(".completedReq").text(filtered.length || 0);
    }
});

$("#returnReq").on("change", function () {
    $(this).parent().find(".customdate").remove();
    const val = $(this).val();
    if (parseInt(val) === 4) {
        currentCustomSection = "returnReq";
        $('#customDateModal').modal('show');
    } else if ($(this).val() == 0) {
        $(".returnReq").text(returnReqForCount.length || 0);
    }
    else {
        const days = getDaysFromDropdown(val);
        const filtered = filterByDays(returnReqForCount, days);
        console.log("filtered", filtered);
        $(".returnReq").text(filtered.length || 0);
    }
});

var totalCasesCount = [];
var openCasesCount = [];
var resolveCaseCount = [];
let currentSection = "";
function getGRCCardsCount(requests) {
    totalCasesCount = requests;
    openCasesCount = requests.filter(item => item.al_finalrequeststatus == 2);
    resolveCaseCount = requests.filter(item => item.al_requeststatus == 6);
    var totalCase = requests.length;
    var openCases = requests.filter(item => item.al_finalrequeststatus == 2)?.length;
    var resolveCases = requests.filter(item => item.al_requeststatus == 6)?.length;
    $(".totalCase").text(totalCase);
    $(".openCases").text(openCases);
    $(".resolveCases").text(resolveCases);
}

$("#totalCase").on("change", function () {
    $(this).parent().find(".customdate").remove();
    const val = $(this).val();
    if (parseInt(val) === 4) {
        currentSection = "totalCase";
        $('#customDateModal').modal('show');
    }
    else if ($(this).val() == 0) {
        $(".totalCase").text(totalCasesCount.length || 0);
    } else {
        const days = getDaysFromDropdown(val);
        const filtered = filterByDays(totalCasesCount, days);
        $(".totalCase").text(filtered.length || 0);
    }
});

$("#openCases").on("change", function () {
    $(this).parent().find(".customdate").remove();
    const val = $(this).val();
    if (parseInt(val) === 4) {
        currentSection = "openCases";
        $('#customDateModal').modal('show');
    } else if ($(this).val() == 0) {
        $(".openCases").text(resolveCaseCount.length || 0);
    }
    else {
        const days = getDaysFromDropdown(val);
        const filtered = filterByDays(resolveCaseCount, days);
        console.log("filtered", filtered);
        $(".openCases").text(filtered.length || 0);
    }
});

$("#resolveCases").on("change", function () {
    $(this).parent().find(".customdate").remove();
    const val = $(this).val();
    if (parseInt(val) === 4) {
        currentSection = "openCases";
        $('#customDateModal').modal('show');
    } else if ($(this).val() == 0) {
        $(".resolveCases").text(resolveCaseCount.length || 0);
    }
    else {
        const days = getDaysFromDropdown(val);
        const filtered = filterByDays(resolveCaseCount, days);
        console.log("filtered", filtered);
        $(".resolveCases").text(filtered.length || 0);
    }
});

function getDaysFromDropdown(val) {
    switch (parseInt(val)) {
        case 1: return 7;
        case 2: return 15;
        case 3: return 30;
        default: return 30;
    }
}

function parseDate(dateStr) {
    return new Date(dateStr); // ISO format or recognizable string
}

// General filtering based on number of days
function filterByDays(requests, days) {
    const now = new Date();
    const pastDate = new Date();
    pastDate.setDate(now.getDate() - days);

    return requests.filter(item => {
        const reqDate = parseDate(new Date(item.createdon));
        return reqDate >= pastDate && reqDate <= now;
    });
}

$("#applyDateFilter").on("click", function () {
    if ($("#customDateForm").valid()) {
        const start = $('#startDate').val();
        const end = $('#endDate').val();

        if (!start || !end || new Date(start) > new Date(end)) {
            alert("Please select a valid date range.");
            return;
        }

        let filtered = [];

        if (currentCustomSection === "completedReq") {
            filtered = filterByCustomDateRange(completedForCount, start, end);
            $("#completedReq").parent().append(`<div class="customdate">Custom Date(${start}'-'${end})</div>`);
            $(".completedReq").text(filtered.length || 0);
        } else if (currentCustomSection === "returnReq") {
            filtered = filterByCustomDateRange(returnReqForCount, start, end);
            $("#returnReq").parent().append(`<div class="customdate">Custom Date(${start} ${end})</div>`);
            $(".returnReq").text(filtered.length || 0);
        }
        else if (currentCustomSection === "totalReq") {
            filtered = filterByCustomDateRange(totalForCount, start, end);
            $("#totalReq").parent().append(`<div class="customdate">Custom Date(${start} ${end})</div>`);
            $(".totalReq").text(filtered.length || 0);
        }
        else if (currentCustomSection === "pendingReq") {
            filtered = filterByCustomDateRange(pendingForCount, start, end);
            $("#pendingReq").parent().append(`<div class="customdate">Custom Date(${start} ${end})</div>`);
            $(".pendingReq").text(filtered.length || 0);
        }
        $('#customDateModal').modal('hide');
    }
});

function filterByCustomDateRange(requests, startDate, endDate) {
    const start = new Date(startDate);
    const end = new Date(endDate);
    start.setHours(0, 0, 0, 0);
    end.setHours(0, 0, 0, 0);
    return requests.filter(item => {
        const reqDate = parseDate(new Date(item.createdon));
        reqDate.setHours(0, 0, 0, 0);
        return reqDate >= start && reqDate <= end;
    });
}

function getFilterOptions() {
    getLegalEntityOptions();
}

function getLegalEntityOptions() {
    webapi.safeAjax({
        type: "GET",
        url: "/_api/cdm_companies?$select=cdm_companycode,cdm_name",
        contentType: "application/json",
        headers: {
            "Prefer": "odata.include-annotations=*"
        },
        success: function (data, textStatus, xhr) {
            var results = data;
            console.log(results);
            // $("#legalEntityFilter").empty().append('<option value="">Select</option>');
            for (var i = 0; i < results.value.length; i++) {
                var result = results.value[i];
                // Columns
                var cdm_companyid = result["cdm_companyid"]; // Guid
                var cdm_companycode = result["cdm_companycode"]; // Text
                var cdm_name = result["cdm_name"]; // Text
                let label = `${cdm_companycode}-${cdm_name}`;
                $("#legalEntityFilter").append(`<option value="${label}">${label}</option>`);
            }
            getAssignToOptions();
        },
        error: function (xhr, textStatus, errorThrown) {
            console.log(xhr);
        }
    });
}

function getAssignToOptions() {
    // Example: multiple team values
    const alTeamTypes = [1, 2, 3, 4, 5, 6, 7, 8, 9, 21, 22, 23, 24, 25]; // OptionSet values for al_teamtype (23 = Category Manager, 24 = SVP Procurement, 25 = Finance)
    //const agUVPTeamTypes = [11, 12, 13, 14, 15, 16, 17, 18]; // OptionSet values for ag_uvpteamtype

    // Build OR conditions for each field
    const alTeamTypeFilter = alTeamTypes
        .map((v) => `al_teamtype eq ${v}`)
        .join(" or ");
    // const agUVPTeamTypeFilter = agUVPTeamTypes
    //     .map((v) => `ag_uvpteamtype eq ${v}`)
    //     .join(" or ");

    // Combine with AND so both fields match
    const filterQuery = `${alTeamTypeFilter}`;

    webapi.safeAjax({
        type: "GET",
        url: `/_api/contacts?$select=contactid,fullname&$filter=${filterQuery}`,
        contentType: "application/json",
        headers: {
            Prefer: "odata.include-annotations=*",
        },
        success: function (data) {
            // $("#assignToFilter").empty().append('<option value="">Select</option>');
            data.value.forEach(function (result) {
                $("#assignToFilter").append(
                    `<option value="${result.fullname}">${result.fullname}</option>`
                );
            });
            console.log(data);
        },
        error: function (xhr) {
            console.log(xhr);
        },
    });
}

//#region Re Assignmet Functionality
$("#assignBtn").on("click", async function () {
    const checkedBoxes = $(".reassignChkBox:checked");
    if (checkedBoxes.length === 0) {
        showErrorModal("Please select at least one request to assign.");
        return;
    }
    try {
        $(".custom-loader").css("display", "flex");
        if (currentUserTeamType == 3) {   //For PSSC Team
            let psscMemberList = await GetEntityList("contacts", `(contactid ne ${loggedInUserId} and _al_psscmanager_value eq ${userWithTeam._al_psscmanager_value} and al_teamtype eq 3 and adx_identity_logonenabled eq true)`, "firstname,fullname,lastname,adx_identity_username");

            // Populate the dropdown
            let dropdown = $('#MutliassigneeSelect');
            dropdown.empty(); // Clear existing options
            dropdown.append(`<option value="" selected="" disabled="">Select Assignee</option>`); // Default option
            if (psscMemberList && psscMemberList.length > 0) {
                psscMemberList.forEach(member => {
                    dropdown.append(`<option value="${member.contactid}">${member.fullname}</option>`);
                });
            } else {
                dropdown.append(`<option value="">No members available</option>`);
            }

            $('#MultiAssignModal').modal('show');       // Show modal
        } else if (currentUserTeamType == 9) {  //For FSSC Team
            let fsscMemberList = await GetEntityList("contacts", `(contactid ne ${loggedInUserId} and _al_fsscmanager_value eq ${userWithTeam._al_fsscmanager_value} and al_teamtype eq 9 and adx_identity_logonenabled eq true)`, "firstname,fullname,lastname,adx_identity_username");

            // Populate the dropdown
            let dropdown = $('#MutliassigneeSelect');
            dropdown.empty(); // Clear existing options
            dropdown.append(`<option value="" selected="" disabled="">Select Assignee</option>`); // Default option
            if (fsscMemberList && fsscMemberList.length > 0) {
                fsscMemberList.forEach(member => {
                    dropdown.append(`<option value="${member.contactid}">${member.fullname}</option>`);
                });
            } else {
                dropdown.append(`<option value="">No members available</option>`);
            }
            $('#MultiAssignModal').modal('show');       // Show modal
        }
    } catch (error) {
        console.log(error);
        showErrorModal(error);
    } finally {
        $(".custom-loader").css("display", "none");
    }
})
async function openAssignModal(assignerId, requestId, approvalId, referenceNumber, vendorName, legalEntityLabel, assignedTo, finalRequestStatusFormatted, createdon, priority, changetype, requestStatus) {
    $('#requestIdHidden').val(requestId);  // Store requestId in hidden input
    $('#approvalIdHidden').val(approvalId);
    $("#currentUserTeam").val(7); //Pssc Manager
    $("#modalReferenceNumber").text(referenceNumber);
    if (vendorName == 'null') {
        $("#modalVendorName").text("N/A");
    } else {
        $("#modalVendorName").text(vendorName);
    }

    if (legalEntityLabel == 'null') {
        $("#modalLegalEntityLabel").text("N/A");
    } else {
        $("#modalLegalEntityLabel").text(legalEntityLabel);
    }
    $("#modalAssignedTo").text(assignedTo);
    $("#modalFinalStatus").text(finalRequestStatusFormatted);
    $("#modalCreatedOn").text(createdon);

    $("#modalApprovalStatus").text(requestStatus);
    if (changetype == 'undefined') {
        $("#modalChangeType").text("N/A")
    } else {
        $("#modalChangeType").text(changetype)
    }
    if (priority == 'undefined') {
        $('.priority-select').html(`<option selected>N/A</option>`);
    }
    else {
        $('.priority-select').html(`<option selected>${priority}</option>`);
    }

    $('#assignModal').validate().resetForm();//reset error
    $('#assignModal').modal('show');// Show modal

    let psscMemberList = await GetEntityList("contacts", `(contactid ne ${assignerId} and al_teamtype eq 3 and _al_psscmanager_value eq ${$("#userIdHidden").html()})`, "firstname,fullname,lastname,adx_identity_username");
    // Populate the dropdown
    let dropdown = $('#assigneeSelect');
    dropdown.empty(); // Clear existing options
    dropdown.append(`<option value="" selected="" disabled="">Assign by email/name</option>`); // Default option

    if (psscMemberList && psscMemberList.length > 0) {
        psscMemberList.forEach(member => {
            dropdown.append(`<option value="${member.contactid}">${member.fullname}</option>`);
        });
    } else {
        dropdown.append(`<option value="">No members available</option>`);
    }
}
async function openAssignModalForFSSC(assignerId, requestId, approvalId, referenceNumber, vendorName, legalEntityLabel, assignedTo, finalRequestStatusFormatted, createdon, priority, changetype, requestStatus) {
    $('#requestIdHidden').val(requestId);  // Store requestId in hidden input
    $('#approvalIdHidden').val(approvalId);
    $("#currentUserTeam").val(8); //FSSC Line Managwer
    $("#modalReferenceNumber").text(referenceNumber);
    if (vendorName == 'null') {
        $("#modalVendorName").text("-");
    } else {
        $("#modalVendorName").text(vendorName);
    }

    if (legalEntityLabel == 'null') {
        $("#modalLegalEntityLabel").text("-");
    } else {
        $("#modalLegalEntityLabel").text(legalEntityLabel);
    }
    $("#modalAssignedTo").text(assignedTo);
    $("#modalFinalStatus").text(finalRequestStatusFormatted);
    $("#modalCreatedOn").text(createdon);

    $("#modalApprovalStatus").text(requestStatus);
    if (changetype == 'undefined') {
        $("#modalChangeType").text("N/A")
    } else {
        $("#modalChangeType").text(changetype)
    }
    if (priority == 'undefined') {
        $('.priority-select').html(`<option selected>N/A</option>`);
    }
    else {
        $('.priority-select').html(`<option selected>${priority}</option>`);
    }

    $('#assignModal').validate().resetForm();//reset error
    $('#assignModal').modal('show');// Show modal

    let fsscMemberList = await GetEntityList("contacts", `(contactid ne ${assignerId} and al_teamtype eq 9 and _al_fsscmanager_value eq ${$("#userIdHidden").html()})`, "firstname,fullname,lastname,adx_identity_username");
    // Populate the dropdown
    let dropdown = $('#assigneeSelect');
    dropdown.empty(); // Clear existing options
    dropdown.append(`<option value="" selected="" disabled="">Assign by email/name</option>`); // Default option

    if (fsscMemberList && fsscMemberList.length > 0) {
        fsscMemberList.forEach(member => {
            dropdown.append(`<option value="${member.contactid}">${member.fullname}</option>`);
        });
    } else {
        dropdown.append(`<option value="">No members available</option>`);
    }
}
$("#assignForm").validate({
    rules: {
    },
    messages: {},
    errorPlacement: function (error, element) {
        if (element.is(":checkbox")) {
        } else if (element.is(":radio")) {
            error.appendTo(element.parents(".form-group"));
        } else {
            error.insertAfter(element);
        }
    },
    submitHandler: async function (form, event) {
        event.preventDefault();
        try {
            $(".custom-loader").css("display", "flex");
            let approvalId = $('#approvalIdHidden').val();
            let updateAppReq = {};
            updateAppReq["al_RequestRaisedBy@odata.bind"] = `/contacts(${contactIdGlobal})`;
            updateAppReq["al_PSSCManager@odata.bind"] = `/contacts(${$("#ManagerIDHidden").val()})`;
            updateAppReq["al_ReAssignTo@odata.bind"] = `/contacts(${$('#assigneeSelect').val()})`;
            if ($("#currentUserTeam").val() == 3) {  //PSSC Member
                updateAppReq.al_reassignaction = 1; //Create By PSSC
                updateAppReq.al_teamtakingaction = 3 //PSSC Member
                updateAppReq.al_requestraisedbyteamtype = 3;//PSSC
                let isReqUpdate = await updateEntityRecord(updateAppReq, "al_approvalrequests", approvalId);
                if (isReqUpdate) {
                    $('#assignModal').modal('hide');
                    setTimeout(() => {
                        $('.custom-loader').css("display", "none"); // First, hide the loader
                        setTimeout(() => {
                            location.reload(); // Then reload the page after a slight delay
                        }, 500); // Give some time for loader to disappear before reloading
                    }, 1000);
                } else {
                    showErrorModal("Request is not Assign.")
                }
            } else if ($("#currentUserTeam").val() == 9) {    //FSSC Team Member
                updateAppReq.al_reassignaction = 1; //Create By FSSC Team
                updateAppReq.al_teamtakingaction = 9 //FSSC Member
                updateAppReq.al_requestraisedbyteamtype = 9;//FSSC Team
                let isReqUpdate = await updateEntityRecord(updateAppReq, "al_approvalrequests", approvalId);
                if (isReqUpdate) {
                    $('#assignModal').modal('hide');
                    setTimeout(() => {
                        $('.custom-loader').css("display", "none"); // First, hide the loader
                        setTimeout(() => {
                            location.reload(); // Then reload the page after a slight delay
                        }, 500); // Give some time for loader to disappear before reloading
                    }, 1000);
                } else {
                    showErrorModal("Request is not Assign.")
                }
            }
            else if (currentUserTeamType == 8) { //FSSC Line Manager
                let approvalId = $('#approvalIdHidden').val();
                let updateApprovalRecord = {};
                updateApprovalRecord["al_RequestRaisedBy@odata.bind"] = `/contacts(${contactIdGlobal})`;
                updateApprovalRecord["al_ReAssignTo@odata.bind"] = `/contacts(${$('#assigneeSelect').val()})`;
                updateApprovalRecord["al_PSSCManager@odata.bind"] = `/contacts(${contactIdGlobal})`;
                updateApprovalRecord.al_reassignaction = 4; // Re Assign By FSSC Manager
                updateApprovalRecord.al_requestraisedbyteamtype = 8; // FSSC Manager
                updateApprovalRecord.al_teamtakingaction = 8 //FSSC Manager
                let isReqUpdate = await updateEntityRecord(updateApprovalRecord, "al_approvalrequests", approvalId);
                if (isReqUpdate) {
                    $('#assignModal').modal('hide');
                    setTimeout(() => {
                        $('.custom-loader').css("display", "none"); // First, hide the loader
                        setTimeout(() => {
                            location.reload(); // Then reload the page after a slight delay
                        }, 500); // Give some time for loader to disappear before reloading
                    }, 1000);
                } else {
                    showErrorModal("Request is not Assign.")
                }
            }
            else {
                let approvalId = $('#approvalIdHidden').val();
                let updateApprovalRecord = {};
                updateApprovalRecord["al_RequestRaisedBy@odata.bind"] = `/contacts(${contactIdGlobal})`;
                updateApprovalRecord["al_ReAssignTo@odata.bind"] = `/contacts(${$('#assigneeSelect').val()})`;
                updateApprovalRecord["al_PSSCManager@odata.bind"] = `/contacts(${contactIdGlobal})`;
                updateApprovalRecord.al_reassignaction = 4; // Re Assign By PSSC Manager
                updateApprovalRecord.al_requestraisedbyteamtype = 7; // PSSC Manager
                updateApprovalRecord.al_teamtakingaction = 7 //PSSC Manager
                let isReqUpdate = await updateEntityRecord(updateApprovalRecord, "al_approvalrequests", approvalId);
                if (isReqUpdate) {
                    $('#assignModal').modal('hide');
                    setTimeout(() => {
                        $('.custom-loader').css("display", "none"); // First, hide the loader
                        setTimeout(() => {
                            location.reload(); // Then reload the page after a slight delay
                        }, 500); // Give some time for loader to disappear before reloading
                    }, 1000);
                } else {
                    showErrorModal("Request is not Assign.")
                }
            }
        } catch (error) {
            showErrorModal("Error in Update Request")
        }
    }
});
// const checkedBoxes = $(".reassignChkBox:checked");
let selectedApprovalIds = new Set();
$(document).on('change', '.reassignChkBox', function () {
    const approvalId = $(this).attr("data-id");

    if (this.checked) {
        selectedApprovalIds.add(approvalId);
    } else {
        selectedApprovalIds.delete(approvalId);
    }
});

$("#MultiAssignForm").validate({
    rules: {
        MutliassigneeSelect: {
            required: true,
        },
    },
    messages: {
        MutliassigneeSelect: {
            required: "Please select an assignee",
        },
    },
    submitHandler: async function (form, event) {
        event.preventDefault();

        $(".custom-loader").css("display", "flex");
        if (selectedApprovalIds.length === 0) {
            alert("Please select at least one request to assign.");
            $(".custom-loader").css("display", "none");
            return false;
        }
        try {
            const updatePromises = Array.from(selectedApprovalIds).map(function (approvalId) {
                const data = {};
                data["al_RequestRaisedBy@odata.bind"] = `/contacts(${contactIdGlobal})`;
                if (currentUserTeamType == 3)  //For PSSC Team
                    data["al_PSSCManager@odata.bind"] = `/contacts(${userWithTeam._al_psscmanager_value})`;
                else if (currentUserTeamType == 9)  //For FSSC Team
                    data["al_PSSCManager@odata.bind"] = `/contacts(${userWithTeam._al_fsscmanager_value})`;

                data["al_ReAssignTo@odata.bind"] = `/contacts(${$("#MutliassigneeSelect").val()})`;
                data.al_reassignaction = 1; //Create By PSSC
                data.al_teamtakingaction = currentUserTeamType; // PSSC Team or FSSC Team
                data.al_requestraisedbyteamtype = currentUserTeamType; //PSSC Team or FSSC Team
                return updateEntityRecord(data, "al_approvalrequests", approvalId);
            });

            const results = await Promise.all(updatePromises);
            const allSucceeded = results.every((r) => r === true);
            if (allSucceeded) {
                $("#MultiAssignModal").modal("hide");
                setTimeout(() => {
                    $(".custom-loader").css("display", "none");
                    setTimeout(() => {
                        location.reload();
                    }, 500);
                }, 1000);
            } else {
                showErrorModal("Some records were not successfully assigned.");
            }
        } catch (error) {
            console.error("Assignment failed:", error);
            alert("An error occurred while assigning. Please try again.");
            $(".custom-loader").css("display", "none");
        } finally {
            $(".custom-loader").css("display", "none");
            selectedApprovalIds.clear();
            $(".reassignChkBox").prop("checked", false); // optional: uncheck all
        }
        return false;
    },
});

//PSSC Member Assign Request to another PSSC Member
async function openPsscAssignModal(psscManagerId, requestId, approvalID, referenceNumber, vendorName, legalEntityLabel, assignedTo, finalRequestStatusFormatted, createdon, priority, changetype, requestStatus) {
    $('#requestIdHidden').val(requestId);  // Store requestId in hidden input
    $("#approvalIdHidden").val(approvalID);
    $("#ManagerIDHidden").val(psscManagerId);
    let loggedInUserId = $("#userIdHidden").html();
    $("#currentUserTeam").val(3);//If this is 3 then pssc assign to another pssc

    $("#modalReferenceNumber").text(referenceNumber);
    if (vendorName == 'null') {
        $("#modalVendorName").text("N/A");
    } else {
        $("#modalVendorName").text(vendorName);
    }

    if (legalEntityLabel == 'null') {
        $("#modalLegalEntityLabel").text("N/A");
    } else {
        $("#modalLegalEntityLabel").text(legalEntityLabel);
    }
    $("#modalAssignedTo").text(assignedTo);
    $("#modalFinalStatus").text(finalRequestStatusFormatted);
    $("#modalCreatedOn").text(createdon);

    $("#modalApprovalStatus").text(requestStatus);
    if (changetype == 'undefined') {
        $("#modalChangeType").text("N/A")
    } else {
        $("#modalChangeType").text(changetype)
    }
    if (priority == 'undefined') {
        $('.priority-select').html(`<option selected>N/A</option>`);
    }
    else {
        $('.priority-select').html(`<option selected>${priority}</option>`);
    }

    let psscMemberList = await GetEntityList("contacts", `(contactid ne ${loggedInUserId} and _al_psscmanager_value eq ${psscManagerId} and al_teamtype eq 3 and adx_identity_logonenabled eq true)`, "firstname,fullname,lastname,adx_identity_username");

    // Populate the dropdown
    let dropdown = $('#assigneeSelect');
    dropdown.empty(); // Clear existing options
    dropdown.append(`<option value="" selected="" disabled="">Assign by email/name</option>`); // Default option
    if (psscMemberList && psscMemberList.length > 0) {
        psscMemberList.forEach(member => {
            dropdown.append(`<option value="${member.contactid}">${member.fullname}</option>`);
        });
    } else {
        dropdown.append(`<option value="">No members available</option>`);
    }

    $('#assignModal').modal('show');       // Show modal
}
//FSSC Member Assign Request to another FSSC Member
async function openFSSCAssignModal(FsscManagerId, requestId, approvalID, referenceNumber, vendorName, legalEntityLabel, assignedTo, finalRequestStatusFormatted, createdon, priority, changetype, requestStatus) {
    $('#requestIdHidden').val(requestId);  // Store requestId in hidden input
    $("#approvalIdHidden").val(approvalID);
    $("#ManagerIDHidden").val(FsscManagerId);
    let loggedInUserId = $("#userIdHidden").html();
    $("#currentUserTeam").val(9);//If this is 9 then FSSC assign to another FSSC

    $("#modalReferenceNumber").text(referenceNumber);
    if (vendorName == 'null') {
        $("#modalVendorName").text("N/A");
    } else {
        $("#modalVendorName").text(vendorName);
    }

    if (legalEntityLabel == 'null') {
        $("#modalLegalEntityLabel").text("N/A");
    } else {
        $("#modalLegalEntityLabel").text(legalEntityLabel);
    }

    $("#modalAssignedTo").text(assignedTo);
    $("#modalFinalStatus").text(finalRequestStatusFormatted);
    $("#modalCreatedOn").text(createdon);

    $("#modalApprovalStatus").text(requestStatus);
    if (changetype == 'undefined') {
        $("#modalChangeType").text("N/A")
    } else {
        $("#modalChangeType").text(changetype)
    }
    if (priority == 'undefined') {
        $('.priority-select').html(`<option selected>N/A</option>`);
    }
    else {
        $('.priority-select').html(`<option selected>${priority}</option>`);
    }

    let fsscMemberList = await GetEntityList("contacts", `(contactid ne ${loggedInUserId} and _al_fsscmanager_value eq ${FsscManagerId} and al_teamtype eq 9 and adx_identity_logonenabled eq true)`, "firstname,fullname,lastname,adx_identity_username");

    // Populate the dropdown
    let dropdown = $('#assigneeSelect');
    dropdown.empty(); // Clear existing options
    dropdown.append(`<option value="" selected="" disabled="">Select Assignee</option>`); // Default option
    if (fsscMemberList && fsscMemberList.length > 0) {
        fsscMemberList.forEach(member => {
            dropdown.append(`<option value="${member.contactid}">${member.fullname}</option>`);
        });
    } else {
        dropdown.append(`<option value="">No members available</option>`);
    }

    $('#assignModal').modal('show');       // Show modal
}
//PSSC Manager or FSSC Manager
async function openPsscModal(id, name, requestId, approvalId, referenceNumber, vendorName, legalEntityLabel, assignedTo, finalRequestStatusFormatted, createdon) {
    $('#psscRequestIdHidden').val(requestId);
    $('#ApprovalIdHidden').val(approvalId);

    let dropdown = $('#psscAssigneeSelect');
    dropdown.empty();
    dropdown.append(`<option value="${id}" selected>${name}</option>`);

    $("#modalPsscReferenceNumber").text(referenceNumber);
    if (vendorName == null) {
        $("#modalPsscVendorName").text("-");
    } else {
        $("#modalPsscVendorName").text(vendorName);
    }

    if (legalEntityLabel == null) {
        $("#modalPsscLegalEntityLabel").text("-");
    } else {
        $("#modalPsscLegalEntityLabel").text(legalEntityLabel);
    }
    $("#modalPsscAssignedTo").text(assignedTo);
    $("#modalPsscFinalStatus").text(finalRequestStatusFormatted);
    $("#modalPsscCreatedOn").text(createdon);

    $('#psscModal').modal('show');       // Show modal
}

//Re Assign ticket Approved by PSSC or FSSC
$("#psscAssignApprovebtn").click(async function () {
    try {
        $(".custom-loader").css("display", "flex");

        let approvalId = $('#ApprovalIdHidden').val();
        // al_approver,al_reassignaction
        let updateApprovalRecord = {};
        updateApprovalRecord["al_Approver@odata.bind"] = `/contacts(${contactIdGlobal})`;
        updateApprovalRecord.al_reassignaction = 2; // Approved
        updateApprovalRecord.al_teamtakingaction = currentUserTeamType;
        let isReqUpdate = await updateEntityRecord(updateApprovalRecord, "al_approvalrequests", approvalId);

        if (isReqUpdate) {
            $('#psscModal').modal('hide');
            setTimeout(() => {
                $('.custom-loader').css("display", "none"); // First, hide the loader
                setTimeout(() => {
                    location.reload(); // Then reload the page after a slight delay
                }, 500); // Give some time for loader to disappear before reloading
            }, 1000);
        } else {
            showErrorModal("Request is not Assign.")
        }
    } catch (error) {
        $('.custom-loader').css("display", "none");
        showErrorModal("Error in Update Request")
    }
});
//Re Assign ticket Rejected by PSSC or FSSC
$("#psscAssignRejectbtn").click(async function () {
    try {
        $(".custom-loader").css("display", "flex");

        let approvalId = $('#ApprovalIdHidden').val();
        // al_approver,al_reassignaction
        let updateApprovalRecord = {};
        updateApprovalRecord["al_Approver@odata.bind"] = `/contacts(${contactIdGlobal})`;
        updateApprovalRecord.al_reassignaction = 3; // Reject
        updateApprovalRecord.al_teamtakingaction = currentUserTeamType;
        let isReqUpdate = await updateEntityRecord(updateApprovalRecord, "al_approvalrequests", approvalId);

        if (isReqUpdate) {
            $('#psscModal').modal('hide');
            setTimeout(() => {
                $('.custom-loader').css("display", "none"); // First, hide the loader
                setTimeout(() => {
                    location.reload(); // Then reload the page after a slight delay
                }, 500); // Give some time for loader to disappear before reloading
            }, 1000);
        } else {
            showErrorModal("Request is not Assign.")
        }
    } catch (error) {
        $('.custom-loader').css("display", "none");
        showErrorModal("Error in Update Request")
    }
});
//#endregion
async function GetUserWithTeam(userId) {
    let userWithTeam = null;
    try {
        let userArr = await GetEntityList("contacts", "contactid eq " + userId, "al_teamtype,contactid,emailaddress1,_al_team_value,fullname,_al_delegator_value,al_isonleave,_al_fsscmanager_value,_al_psscmanager_value");
        userWithTeam = Array.isArray(userArr) ? userArr[0] : null;
    } catch (error) {
        console.log(error);
    }
    return userWithTeam;
}

async function GetP1Requests(userId) {
    let todayDateTime = new Date().toISOString();
    // let reqList = await GetEntityList("al_requests", `al_priorityp1 eq 0 and createdon lt ${todayDateTime}`, "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendorname,al_requeststatus,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_reassigndetail,al_subrequesttype,al_vendormultiplelabel,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority", "modifiedon desc");
    let reqList = await GetEntityList(
        "al_requests",
        `(_al_requestedby_value eq ${userId} or _al_primarycontact_value eq ${userId}) and al_priority eq 1 and createdon lt ${todayDateTime}`,
        "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendorname,al_requeststatus,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_reassigndetail,al_subrequesttype,al_vendormultiplelabel,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority",
        "modifiedon desc",
        "al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon),al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)"
    );

    return reqList;
}
async function GetBuTeamRequests(userId, type) {
    let todayDateTime = new Date().toISOString();

    if (type === 'Pending') {
        // Filter logic: al_requeststatus == 1 || al_requeststatus == 16 || al_requeststatus == 41 ||
        // (al_requeststatus == 5 && _al_primarycontact_value == userId) ||
        // (_al_primarycontact_value == userId && al_requeststatus != 8)
        let filter = `
            _al_requestedby_value eq ${userId}
            and createdon lt ${todayDateTime}
            and (
                al_requeststatus eq 1
                or al_requeststatus eq 16
                or al_requeststatus eq 41
                or (al_requeststatus eq 5 and _al_primarycontact_value eq ${userId})
                or (_al_primarycontact_value eq ${userId} and al_requeststatus ne 8)
            )
        `;

        let reqList = await GetEntityList(
            "al_requests",
            filter.replace(/\s+/g, " ").trim(),
            "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendorname,al_requeststatus,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_reassigndetail,al_subrequesttype,al_vendormultiplelabel,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority",
            "modifiedon desc",
            "al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon),al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)"
        );

        return reqList;
    }
    else if (type === "All") {
        let reqList = await GetEntityList(
            "al_requests",
            `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime}`,
            "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendorname,al_requeststatus,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_reassigndetail,al_subrequesttype,al_vendormultiplelabel,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority",
            "modifiedon desc",
            "al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon),al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)"
        );
        return reqList;
    }
    else if (type == "Returned") {
        let reqList = await GetEntityList(
            "al_requests",
            `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime} and al_requeststatus eq 5`,
            "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendorname,al_requeststatus,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_reassigndetail,al_subrequesttype,al_vendormultiplelabel,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority",
            "modifiedon desc",
            "al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon),al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)"
        );
        return reqList;
    }
    else if (type == "Resolved") {
        let reqList = await GetEntityList(
            "al_requests",
            `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime} and (al_requeststatus eq 6 or al_requeststatus eq 40)`,
            "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendorname,al_requeststatus,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_reassigndetail,al_subrequesttype,al_vendormultiplelabel,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority",
            "modifiedon desc",
            "al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon),al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)"
        );
        return reqList;
    }
    else if (type == "Rejected") {
        let reqList = await GetEntityList(
            "al_requests",
            `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime} and al_requeststatus eq 4`,
            "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendorname,al_requeststatus,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_reassigndetail,al_subrequesttype,al_vendormultiplelabel,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority",
            "modifiedon desc",
            "al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon),al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)"
        );
        return reqList;
    }
}

async function GetManagerRequests(userId) {
    let todayDateTime = new Date().toISOString();
    var now = new Date();
    var isoDateTime = now.toISOString();
    let approvalRequests = await GetEntityList(
        "al_approvalrequests",
        `_al_primarycontact_value eq ${userId} and _al_request_value ne null and createdon lt ${todayDateTime} and Microsoft.Dynamics.CRM.OnOrBefore(PropertyName='modifiedon',PropertyValue='${isoDateTime}')`,
        `al_approvalrequestid,al_teamtype,_al_primarycontact_value,_al_request_value&$expand=al_Request
        (
        $select=al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,
        al_requeststatus,al_requesttype,_al_currentapprovalrequest_value,al_reassigndetail,al_vendormultiplelabel,al_prnumber,al_vendorname,al_ponumber,_al_requestedby_value,al_subrequesttype,al_finalrequeststatus,al_reassigndetailfssc,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_changetype,al_priority;$expand=al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail),al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon)
        )`,
        "modifiedon desc"
    );
    let requestList = approvalRequests
        .map(element => {
            let request = element.al_Request;
            // Copy UVP vendor request data from approval request to request object
            if (element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest) {
                request.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest = element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest;
            }
            return request;
        }) // Extract only the request object
        .filter(req => req && req.al_requestid) // Ensure it's not null
    let requests = await GetEntityList(
        "al_requests",
        `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime}`,
        "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_requeststatus,al_requesttype,al_eauctionname,_al_requestedby_value,al_prnumber,al_ponumber,al_vendorname,al_subrequesttype,al_finalrequeststatus,_al_primarycontact_value,al_projectnameid,al_changetype,al_priority",
        "modifiedon desc",
        "al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail),al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon)"
    );

    requestList.push(...requests);
    // Remove duplicates based on `al_requestid`
    const uniqueList = Array.from(
        new Map(requestList.map(item => [item.al_requestid, item])).values()
    );

    // Sort by `createdon` in descending order
    uniqueList.sort((a, b) => new Date(b.modifiedon) - new Date(a.modifiedon));

    return uniqueList;
}
//Get All Record For PSSC Team Memeber
async function GetAllRequestForPSSC(userId) {
    let todayDateTime = new Date().toISOString();
    let approvalRequests = await GetEntityList(
        "al_approvalrequests",
        `al_teamtype eq 3 and _al_request_value ne null and createdon lt ${todayDateTime}`,
        `al_approvalrequestid,_al_primarycontact_value,_al_request_value&$expand=al_Request
        (
        $select=al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_vendormultiplelabel,al_psscreassignmembername,al_psscreassignmemberid,
        al_requeststatus,al_requesttype,al_eauctionname,al_projectnameid,al_reassigndetail,al_reassigndetailfssc,_al_requestedby_value,al_vendorname,al_subrequesttype,al_prnumber,al_ponumber,al_finalrequeststatus,_al_primarycontact_value,_al_currentapprovalrequest_value,al_changetype,al_priority;$expand=al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail),al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon)
        )`,
        "modifiedon desc"
    );//change

    let requestList = approvalRequests
        .map(element => {
            let request = element.al_Request;
            // Copy UVP vendor request data from approval request to request object
            if (element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest) {
                request.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest = element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest;
            }
            return request;
        }) // Extract only the request object
        .filter(req => req && req.al_requestid) // Ensure it's not null
    let requests = await GetEntityList("al_requests", `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime} and (al_subrequesttype eq 13 or al_subrequesttype eq 14)`, "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_requeststatus,al_vendorname,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,al_requesttype,al_subrequesttype,al_finalrequeststatus,_al_primarycontact_value", "modifiedon desc", "al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail),al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon)");

    requestList.push(...requests);
    // Remove duplicates based on `al_requestid`
    const uniqueList = Array.from(
        new Map(requestList.map(item => [item.al_requestid, item])).values()
    );

    // Sort by `createdon` in descending order
    uniqueList.sort((a, b) => new Date(b.modifiedon) - new Date(a.modifiedon));

    return uniqueList;
}
//Get All Record for PSSC Manager
async function GetAllRecordForPSSCManager(userId) {
    let todayDateTime = new Date().toISOString();
    //42 for re assign request status
    let approvalRequests = await GetEntityList(
        "al_approvalrequests",
        `(al_teamtype eq 3 and al_Request/al_finalrequeststatus eq 2 and (al_Request/al_requeststatus eq 3 or al_Request/al_requeststatus eq 42) and _al_request_value ne null and 
        al_PrimaryContact/_al_psscmanager_value eq ${userId}) and createdon lt ${todayDateTime}`,
        `al_approvalrequestid,_al_primarycontact_value,_al_request_value,al_teamtype&$expand=al_Request
        (
        $select=al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_psscreassignmemberid,al_psscreassignmembername,
        al_requeststatus,al_vendorname,al_vendormultiplelabel,al_requesttype,al_prnumber,al_ponumber,_al_requestedby_value,_al_currentapprovalrequest_value,al_subrequesttype,al_finalrequeststatus,_al_primarycontact_value,al_eauctionname,al_projectnameid,al_reassigndetailfssc,al_reassigndetail,al_changetype,al_priority;$expand=al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail),al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon)
        )`,
        "modifiedon desc"
    );
    let requestList = approvalRequests
        .map(element => {
            let request = element.al_Request;
            // Copy UVP vendor request data from approval request to request object
            if (element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest) {
                request.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest = element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest;
            }
            return request;
        }) // Extract only the request object
        .filter(req => req && req.al_requestid) // Ensure it's not null

    // Remove duplicates based on `al_requestid`
    const uniqueList = Array.from(
        new Map(requestList.map(item => [item.al_requestid, item])).values()
    );

    // Sort by `createdon` in descending order
    uniqueList.sort((a, b) => new Date(b.modifiedon) - new Date(a.modifiedon));

    return uniqueList;
}

//Get All Record For FSSC Manager
async function GetAllRecordForFSSCManager(userId) {
    //42 for re assign request status
    let todayDateTime = new Date().toISOString();
    let approvalRequests = await GetEntityList(
        "al_approvalrequests",
        `((_al_primarycontact_value eq ${userId} and _al_request_value ne null) or (_al_psscmanager_value eq ${userId} and al_Request/al_requeststatus eq 42) or (al_PrimaryContact/_al_fsscmanager_value eq ${userId} and al_teamtype eq 9 and al_requeststatus eq 1))`,
        `al_approvalrequestid,_al_primarycontact_value,_al_request_value,al_teamtype&$expand=al_Request
        (
        $select=al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_psscreassignmemberid,al_psscreassignmembername,
        al_requeststatus,al_requesttype,al_vendormultiplelabel,al_prnumber,al_ponumber,_al_requestedby_value,al_vendorname,_al_currentapprovalrequest_value,al_subrequesttype,al_finalrequeststatus,_al_primarycontact_value,al_reassigndetail ,al_reassigndetailfssc,al_eauctionname,al_projectnameid,al_changetype,al_priority;$expand=al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail),al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon)
        )`,
        "modifiedon desc"
    );
    let requests = await GetEntityList("al_requests", `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime} and (al_subrequesttype eq 13 or al_subrequesttype eq 14)`, "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_prnumber,al_ponumber,al_requeststatus,al_vendorname,al_requesttype,_al_requestedby_value,al_subrequesttype,al_finalrequeststatus,_al_primarycontact_value", "modifiedon desc", "al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)");

    let requestList = approvalRequests
        .map(element => {
            let request = element.al_Request;
            // Copy UVP vendor request data from approval request to request object
            if (element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest) {
                request.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest = element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest;
            }
            return request;
        }) // Extract only the request object
        .filter(req => req && req.al_requestid) // Ensure it's not null

    requestList.push(...requests);

    // Remove duplicates based on `al_requestid`
    const uniqueList = Array.from(
        new Map(requestList.map(item => [item.al_requestid, item])).values()
    );

    // Sort by `createdon` in descending order
    uniqueList.sort((a, b) => new Date(b.modifiedon) - new Date(a.modifiedon));

    return uniqueList;
}
//Update All Entity Record
async function updateEntityRecord(data, schemaName, recordId = null) {
    try {
        if (recordId) {
            let isUpdated = false;

            await webapi.safeAjax({
                type: "PATCH",
                contentType: "application/json",
                url: `/_api/${schemaName}(${recordId})`,
                data: JSON.stringify(data),
                success: function (data, textStatus, xhr) {
                    console.log("Record updated");
                    isUpdated = true;
                },
                error: function (xhr, textStatus, errorThrown) {
                    console.error(`Error updating ${schemaName}:`, xhr, textStatus, errorThrown);
                    showErrorModal("Failed to update record. Please try again.");
                }
            });
            return isUpdated;
        } else {
            showErrorModal("Request id not found. Please try again.");
        }
    } catch (error) {
        console.error(`Unexpected error in createOrUpdateEntityRecord:`, error);
        showErrorModal("An unexpected error occurred. Please contact support.");
        return null;
    } finally {
        $('.custom-loader').css("display", "none");
    }
}

function showErrorModal(message) {
    $("#errorMessage").text(message || "Something went wrong.");
    $("#errorModal").modal("show");
}

async function GetEntityList(schemaName, filter = null, fieldname, sort = null, expand = null) {
    let list = [];

    if (webapi?.safeAjax) {
        let finalFilter = filter ? `&$filter=${filter}` : "";
        let finalOrderBy = sort ? `&$orderby=${sort}` : "";
        let finalExpand = expand == null ? "" : "&$expand=" + expand;
        let nextLink = `/_api/${schemaName}?$select=${fieldname}${finalFilter}${finalExpand}${finalOrderBy}`;


        while (nextLink) {
            try {
                let data = await webapi.safeAjax({
                    type: "GET",
                    url: nextLink,
                    contentType: "application/json",
                    headers: {
                        "Prefer": "odata.include-annotations=*"
                    }
                });

                list = list.concat(data?.value || []);
                nextLink = data["@odata.nextLink"] || null;

            } catch (error) {
                console.error("Error fetching data:", error);
                // Don't show error modal for SLA settings - it's not critical
                if (schemaName !== "al_slasettings") {
                    $('.custom-loader').css("display", "none");
                    showErrorModal("Error fetching data.");
                }
                return null; // Return null on error
            }
        }
    }

    return list;
}

//Download Approval Request Excel Sheet
$("#approvalRequestExcel").click(async function () {
    $('.custom-loader').css("display", "flex");
    let errorMessage = "";
    try {
        if (currentUserTeamType) {
            let loggedInUserId = $("#userIdHidden").html();
            await $.ajax({
                type: "GET",
                async: false,
                url: PSSC_DownloadApprovalRequest + `&userId=${loggedInUserId}&teamType=${currentUserTeamType}`,
                success: function (response) {
                    downloadExcelFromBase64(response, "Approval Request");
                },
                error: function (error) {
                    errorMessage = error.message;
                    console.log(error);
                }
            });
        }
    } catch (error) {
        errorMessage = error.message;
        $('.custom-loader').css("display", "none");
        showErrorModal(errorMessage);
    }
    finally {
        $('.custom-loader').css("display", "none");
    }
});
//Download Request Excel Sheet
$("#requestExcel").click(async function () {
    $('#requestModal').modal('show');// Show modal
});

$("#requestExcelDownloadBtn").click(async function () {
    $('.custom-loader').css("display", "flex");
    let errorMessage = "";
    try {
        let subReqTypeValue = $("#subReqTypeSelect").val();
        if (currentUserTeamType && subReqTypeValue) {
            let loggedInUserId = $("#userIdHidden").html();
            await $.ajax({
                type: "GET",
                async: false,
                url: PSSC_DownloadRequest + `&userId=${loggedInUserId}&teamType=${currentUserTeamType}&subReqType=${subReqTypeValue}`,
                success: function (response) {
                    downloadExcelFromBase64(response, "Request");
                },
                error: function (error) {
                    errorMessage = error.message;
                    console.log(error);
                }
            });
        }
    } catch (error) {
        errorMessage = error.message;
        $('.custom-loader').css("display", "none");
        showErrorModal(errorMessage);
    }
    finally {
        $('.custom-loader').css("display", "none");
        $('#requestModal').modal('hide');// Show modal
    }
});

function downloadExcelFromBase64(base64, name) {
    const link = document.createElement("a");
    link.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + base64;
    link.download = name + ".xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

async function GetComplianceRequests(userId) {
    let todayDateTime = new Date().toISOString();
    let approvalRequests = await GetEntityList(
        "al_approvalrequests",
        `_al_primarycontact_value eq ${userId} and _al_request_value ne null and createdon lt ${todayDateTime}`,
        `al_approvalrequestid,_al_primarycontact_value,_al_request_value,al_requeststatus&$expand=al_Request
        (
        $select=al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,
        al_requeststatus,al_requesttype,al_prnumber,al_reassigndetail,al_ponumber,_al_currentapprovalrequest_value,_al_requestedby_value,al_subrequesttype,al_finalrequeststatus,_al_primarycontact_value,al_changetype,al_priority;$expand=al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail),al_approvalrequest_Request_al_request($select=al_requeststatus,_al_approvedby_value,al_name,_al_primarycontact_value,al_referencenumber,_al_team_value,al_teamtype,createdon)
        )`,
        "modifiedon desc"
    );
    let requests = await GetEntityList("al_requests", `_al_requestedby_value eq ${userId} and createdon lt ${todayDateTime} and (al_subrequesttype eq 13 or al_subrequesttype eq 14)`, "al_requestid,createdon,modifiedon,al_legalentitydetails,_al_raisedonbehalfof_value,al_referencenumber,al_requeststatus,al_requesttype,_al_requestedby_value,al_subrequesttype,al_finalrequeststatus,_al_primarycontact_value", "modifiedon desc", "al_UVPVendorRequest($select=ag_requestnumber,ag_vendorcode,ag_vendorcontactemail)");

    let requestList = approvalRequests
        .map(element => {
            if (element.al_Request && element.al_Request.al_requestid) {
                let request = {
                    ...element.al_Request,
                    al_approvalReqStatus: element.al_requeststatus // Add new property
                };
                // Copy UVP vendor request data from approval request to request object
                if (element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest) {
                    request.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest = element.ag_uvpvendorrequest_CurrentApprovalRequest_al_approvalrequest;
                }
                return request;
            }
            return null; // Filter later
        })
        .filter(req => req !== null); // Remove nulls

    grcRejectRequestList = Array.from(
        new Map(requestList.filter(x => x.al_approvalReqStatus == 6).map(item => [item.al_requestid, item])).values()
    );

    grcApproveRequestList = Array.from(
        new Map(requestList.filter(x => x.al_approvalReqStatus == 4).map(item => [item.al_requestid, item])).values()
    );

    requestList.push(...requests);
    // Remove duplicates based on `al_requestid`
    const uniqueList = Array.from(
        new Map(requestList.map(item => [item.al_requestid, item])).values()
    );

    // Sort by `createdon` in descending order
    uniqueList.sort((a, b) => new Date(b.modifiedon) - new Date(a.modifiedon));

    return uniqueList;
}

$(document).on("click", "#applyCustomDateFilterBtn", function () {
    if ($('#customDateFilterForm').valid()) {
        let customdate = "Custom Date(" + $("#filterFromDate").val() + " - " + $("#filterToDate").val() + ")";
        $("#dateFilter").siblings(".select2-container").find("#select2-dateFilter-container").text(customdate);
        $("#dateFilter").find('option[value = "custom"]').text(customdate);
        $("#dateFilter").select2();
        $("#dateFilter").siblings(".select2-container").find("#select2-dateFilter-container").attr("title", customdate);
        $("#customDateFilterModal").modal('hide');
    }
});

// Common function to apply filters and update pagination
async function applyFiltersAndUpdatePagination() {
    var table;

    var userWithTeam;

    var teamType = -1;



    // Get current user and team info

    try {

        const loggedInUserId = $("#userIdHidden").html();

        if (loggedInUserId) {

            userWithTeam = await GetUserWithTeam(loggedInUserId);

            if (userWithTeam) {

                teamType = userWithTeam.al_teamtype;

            }

        }

    } catch (error) {

        console.error("Error getting user with team:", error);

    }



    if ($('#pendingReqTable').is(":visible")) {

        table = $('#pendingReqTable').DataTable();

    } else if ($('#grcApproveTable').is(":visible")) {

        table = $('#grcApproveTable').DataTable();

    }

    else if ($('#grcRejectTable').is(":visible")) {

        table = $('#grcRejectTable').DataTable();

    }

    else if ($('#returenedRequestTable').is(":visible")) {

        table = $('#returenedRequestTable').DataTable();

    }

    else if ($('#p1Table').is(":visible")) {

        table = $('#p1Table').DataTable();

    }

    else {

        table = $('#allReqTable').DataTable();

    }



    // Check if we have stored data to filter (for case 1 - BU team)

    if (allRequestDataStorage && allRequestDataStorage.length > 0) {

        // Collect filter criteria

        var requestSubType = $('#requestSubtype').val() || [];

        var legalEntity = $("#legalEntityFilter").val() || [];

        var assignTo = $("#assignToFilter").val() || [];

        var approvalStatus = $('#approvalStatus').val() || [];

        var finalStatus = $("#finalStatus").val() || [];

        var searchText = $('#searchText').val()?.trim() || null;



        // Date filter

        var dateFilter = $('#dateFilter').val();

        var fromDate = $('#filterFromDate').val();

        var toDate = $('#filterToDate').val();

        let minDate = null, maxDate = null;

        if (dateFilter === 'last7') {

            minDate = moment().subtract(7, 'days').format('YYYY-MM-DD');

        } else if (dateFilter === 'last15') {

            minDate = moment().subtract(15, 'days').format('YYYY-MM-DD');

        } else if (dateFilter === 'last30') {

            minDate = moment().subtract(30, 'days').format('YYYY-MM-DD');

        } else if (dateFilter === 'custom') {

            minDate = fromDate || null;

            maxDate = toDate || null;

        }



        // If all filters are empty, reset to show all

        if (!requestSubType.length && !legalEntity.length && !assignTo.length &&

            !approvalStatus.length && !finalStatus.length && !searchText &&

            !minDate && !maxDate) {

            // Reset to original data

            //allPendingRequests = allRequestDataStorage;

            filteredRequestData = null;

            table.clear();

            const firstBatch = allRequestDataStorage.slice(0, RECORDS_PER_BATCH);

            await BindRequest(userWithTeam, firstBatch, table, -1, null, null, false, true);

            // $(".pendingReq").text(allRequestDataStorage.length);

            loadedPendingBatches = 1;

            // Update pagination to reflect original count

            // Update pagination to reflect original count - update immediately
            if (table) {
                // Update pagination settings immediately
                window.updateDataTablesPagination(table, allRequestDataStorage.length);

                // Also update the info text manually to ensure it's visible
                setTimeout(() => {
                    try {
                        if (typeof table.api === 'function') {
                            let api = table.api();
                            if (api) {
                                let pageInfo = api.page.info();
                                let start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1;
                                let end = pageInfo.end;
                                let total = allRequestDataStorage.length;
                                let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                                if ($info.length) {
                                    $info.text(`Showing ${start} to ${end} of ${total} entries`);
                                }
                                // Force a redraw to update pagination numbers
                                api.draw(false);
                            }
                        }
                    } catch (error) {
                        console.warn('Error updating pagination info:', error);
                    }
                }, 150);
            }



            // Update KPI cards with original data counts

            // if (allRequestDataStorage && allRequestDataStorage.length > 0) {
            //     getRequestCount(allRequestDataStorage);
            // }
            // Note: KPI cards should NOT be updated here - they should always show total counts
            // KPI counts are set when data is initially loaded, not when filtering
            return;
        }



        // Apply filters to stored data

        await applyFiltersToStoredData({

            searchText: searchText,

            requestSubType: requestSubType,

            legalEntity: legalEntity,

            assignTo: assignTo,

            approvalStatus: approvalStatus,

            finalStatus: finalStatus,

            minDate: minDate ? moment(minDate) : null,

            maxDate: maxDate ? moment(maxDate) : null

        }, userWithTeam, table, -1);



        return; // Exit early, filtering is done

    }



    // Fallback to original DataTables filtering for other cases


    var requestSubType = $('#requestSubtype').val();

    var legalEntity = $("#legalEntityFilter").val();

    var assignTo = $("#assignToFilter").val();

    var approvalStatus = $('#approvalStatus').val();

    var finalStatus = $("#finalStatus").val();

    // var legalEntity = $("#legalEntityFilter").val() && $("#legalEntityFilter option:selected").text().toLowerCase() !== 'select'

    //     ? $("#legalEntityFilter option:selected").text().toLowerCase()

    //     : null;


    // var assignTo = $("#assignToFilter").val() && $("#assignToFilter option:selected").text().toLowerCase() !== 'select'

    //     ? $("#assignToFilter option:selected").text().toLowerCase()

    //     : null;


    // var approvalStatus = $('#approvalStatus').val() && $('#approvalStatus option:selected').text().toLowerCase() !== 'select'

    //     ? $('#approvalStatus option:selected').text().toLowerCase()

    //     : null;


    // var finalStatus = $("#finalStatus").val() && $("#finalStatus option:selected").text().toLowerCase() !== 'select'

    //     ? $("#finalStatus option:selected").text().toLowerCase()

    //     : null;


    var searchText = $('#searchText').val()?.trim().toLowerCase() || null;


    // Date filter

    var dateFilter = $('#dateFilter').val();

    var fromDate = $('#filterFromDate').val();

    var toDate = $('#filterToDate').val();


    let minDate = null, maxDate = null;

    if (dateFilter === 'last7') {

        minDate = moment().subtract(7, 'days').format('YYYY-MM-DD');

    } else if (dateFilter === 'last15') {

        minDate = moment().subtract(15, 'days').format('YYYY-MM-DD');

    } else if (dateFilter === 'last30') {

        minDate = moment().subtract(30, 'days').format('YYYY-MM-DD');

    } else if (dateFilter === 'custom') {

        minDate = fromDate || null;

        maxDate = toDate || null;

    }


    // Remove previous filters

    $.fn.dataTable.ext.search = [];


    // If all filters are empty, reset to show all

    if (

        !requestSubType && !legalEntity && !assignTo &&

        !approvalStatus && !finalStatus && !searchText &&

        !minDate && !maxDate

    ) {

        table.draw();

        return;

    }


    // Filtering logic

    $.fn.dataTable.ext.search.push(function (settings, data, dataIndex) {


        // let rowAdded = tableElement.row.add({

        //                 DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: firstColumn, 1: referenceNumber, 2: subrequesttypeFormatted, 3: createdOn, 4: vendorName, 5: legalEntityLabel, 6: assignedTo, 7: requestStatus, 8: finalRequestStatusFormatted, 9:  `<div class="d-flex justify-content-center">${psscReAssignBtn}</div>`, 10: prNumber, 11: poNumber,

        //                 12: ProjectID, 13: EAuctionName

        //             });


        let requestDateStr;

        let requestDate;

        let matchSubType;

        let matchLegalEntity;

        let matchAssignTo;

        let matchApprovalStatus;

        let matchFinalStatus;

        let matchSearchText;

        if ($('#pendingReqTable').is(":visible")) {

            requestDateStr = data[4];  // e.g. "7/22/2025 5:32 PM"

            requestDate = moment(requestDateStr, 'M/D/YYYY h:mm A');

            // matchSubType = !requestSubType || data[3].toLowerCase().includes(requestSubType);

            matchSubType = requestSubType.length === 0 || requestSubType.some(subType =>

                // data[3].toLowerCase().includes(subType.toLowerCase())

                data[3].trim().toLowerCase() === subType.trim().toLowerCase()

            );

            // matchLegalEntity = !legalEntity || data[6].toLowerCase().includes(legalEntity);

            matchLegalEntity = legalEntity.length === 0 || legalEntity.some(entity =>

                // data[6].toLowerCase().includes(entity.toLowerCase())

                data[8].trim().toLowerCase() === entity.trim().toLowerCase()

            );


            // matchAssignTo = !assignTo || data[7].toLowerCase().includes(assignTo);

            matchAssignTo = assignTo.length === 0 || assignTo.some(assignee =>

                // data[7].toLowerCase().includes(assignee.toLowerCase())

                data[9].trim().toLowerCase() === assignee.trim().toLowerCase()

            );


            // matchApprovalStatus = !approvalStatus || data[11].toLowerCase().includes(approvalStatus);

            matchApprovalStatus = approvalStatus.length === 0 || approvalStatus.some(approveStatus =>

                // data[11].toLowerCase().includes(approveStatus.toLowerCase())

                data[13].trim().toLowerCase() === approveStatus.trim().toLowerCase()

            );


            // matchFinalStatus = !finalStatus || data[12].toLowerCase().includes(finalStatus);

            matchFinalStatus = finalStatus.length === 0 || finalStatus.some(finalReqStatus =>

                // data[12].toLowerCase().includes(finalReqStatus.toLowerCase())

                data[14].trim().toLowerCase() === finalReqStatus.trim().toLowerCase()


            );


            // Global search in Pending tab: match if text appears in ANY column of the row

            if (!searchText) {

                matchSearchText = true;

            } else {

                const lowerSearch = searchText.toLowerCase();

                matchSearchText = data.some(col =>

                    (col || '').toString().toLowerCase().includes(lowerSearch)

                );

            }

        } else {

            // let rowAdded = tableElement.row.add({

            //             DT_RowAttr: { 'title': reassignDetail, 'data-auctionname': result.al_eauctionname, 'data-projectid': result.al_projectnameid }, __raw: result, 0: referenceNumber, 1: uvpRequestId, 2: subrequesttypeFormatted, 3: createdOn, 4: vendorName, 5: legalEntityLabel, 6: assignedTo, 7: requestedBy, 8: requestStatus, 9: finalRequestStatusFormatted, 10: `<div class="d-flex justify-content-center">${actionBtn}</div>`, 11:  prNumber, 12: poNumber, 13: ProjectID,

            //             14: EAuctionName

            //         });


            requestDateStr = data[3];  // e.g. "7/22/2025 5:32 PM"

            requestDate = moment(requestDateStr, 'M/D/YYYY h:mm A');

            // matchSubType = !requestSubType || data[2].toLowerCase().includes(requestSubType);

            matchSubType = requestSubType.length === 0 || requestSubType.some(subType =>

                data[2].toLowerCase().includes(subType.toLowerCase())


            );

            // matchLegalEntity = !legalEntity || data[5].toLowerCase().includes(legalEntity);

            matchLegalEntity = legalEntity.length === 0 || legalEntity.some(entity =>

                data[7].toLowerCase().includes(entity.toLowerCase()));


            // matchAssignTo = !assignTo || data[6].toLowerCase().includes(assignTo);

            matchAssignTo = assignTo.length === 0 || assignTo.some(assignee =>

                data[8].toLowerCase().includes(assignee.toLowerCase()));

            // matchApprovalStatus = !approvalStatus || data[10].toLowerCase().includes(approvalStatus);

            matchApprovalStatus = approvalStatus.length === 0 || approvalStatus.some(approveStatus =>

                data[12].toLowerCase().includes(approveStatus.toLowerCase()));

            // matchFinalStatus = !finalStatus || data[11].toLowerCase().includes(finalStatus);

            matchFinalStatus = finalStatus.length === 0 || finalStatus.some(finalReqStatus =>

                data[13].toLowerCase().includes(finalReqStatus.toLowerCase()));

            // Global search in other tabs: match if text appears in ANY column of the row

            if (!searchText) {

                matchSearchText = true;

            } else {

                const lowerSearch = searchText.toLowerCase();

                matchSearchText = data.some(col =>

                    (col || '').toString().toLowerCase().includes(lowerSearch)

                );

            }

        }



        let matchDate = true;

        if (minDate || maxDate) {

            const rowDate = requestDate;

            if (minDate && rowDate.isBefore(minDate, 'day')) matchDate = false;

            if (maxDate && rowDate.isAfter(maxDate, 'day')) matchDate = false;

        }
        return matchSubType && matchLegalEntity && matchAssignTo &&

            matchApprovalStatus && matchFinalStatus &&

            matchSearchText && matchDate;

    });
    table.draw();
}

// Function to filter stored data based on filter criteria

async function applyFiltersToStoredData(filterCriteria, userWithTeam, tableElement, teamType) {
    if (!allRequestDataStorage || allRequestDataStorage.length === 0) {
        console.warn("No data stored for filtering");
        return;
    }
    let filteredData = [...allRequestDataStorage]; // Start with all data
    // Apply search text filter
    if (filterCriteria.searchText) {
        const searchLower = filterCriteria.searchText.toLowerCase();
        filteredData = filteredData.filter(result => {
            const referenceNumber = (result["al_referencenumber"] || "").toLowerCase();
            const vendorName = (result["al_vendorname"] || result["al_vendormultiplelabel"] || "").toLowerCase();
            const vendorCode = (result.al_UVPVendorRequest?.ag_vendorcode || "").toLowerCase();
            const vendorEmail = (result.al_UVPVendorRequest?.ag_vendorcontactemail || "").toLowerCase();
            const legalEntity = (result["al_legalentitydetails"] || "").toLowerCase();
            const assignedTo = (result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "").toLowerCase();
            const requestedBy = (result["_al_requestedby_value@OData.Community.Display.V1.FormattedValue"] || "").toLowerCase();
            const requestStatus = (result["al_requeststatus@OData.Community.Display.V1.FormattedValue"] || "").toLowerCase();
            const finalStatus = (result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"] || "").toLowerCase();
            const poNumber = (result.al_ponumber || "").toLowerCase();
            const prNumber = (result.al_prnumber || "").toLowerCase();

            const uvpRequestId = (result.al_UVPVendorRequest?.ag_requestnumber || "").toLowerCase();

            const subRequestType = (result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"] || "").toLowerCase();



            return referenceNumber.includes(searchLower) ||

                vendorName.includes(searchLower) ||

                vendorCode.includes(searchLower) ||

                vendorEmail.includes(searchLower) ||

                legalEntity.includes(searchLower) ||

                assignedTo.includes(searchLower) ||

                requestedBy.includes(searchLower) ||

                requestStatus.includes(searchLower) ||

                finalStatus.includes(searchLower) ||

                poNumber.includes(searchLower) ||

                prNumber.includes(searchLower) ||

                uvpRequestId.includes(searchLower) ||

                subRequestType.includes(searchLower);

        });

    }



    // Apply request sub type filter

    if (filterCriteria.requestSubType && filterCriteria.requestSubType.length > 0) {

        filteredData = filteredData.filter(result => {

            const subTypeFormatted = (result["al_subrequesttype@OData.Community.Display.V1.FormattedValue"] || "").trim().toLowerCase();

            return filterCriteria.requestSubType.some(subType =>

                subTypeFormatted === subType.trim().toLowerCase()

            );

        });

    }



    // Apply legal entity filter

    if (filterCriteria.legalEntity && filterCriteria.legalEntity.length > 0) {

        filteredData = filteredData.filter(result => {

            const legalEntity = (result["al_legalentitydetails"] || "").trim().toLowerCase();

            return filterCriteria.legalEntity.some(entity =>

                legalEntity === entity.trim().toLowerCase()

            );

        });

    }



    // Apply assign to filter

    if (filterCriteria.assignTo && filterCriteria.assignTo.length > 0) {

        filteredData = filteredData.filter(result => {

            const assignedTo = (result["_al_primarycontact_value@OData.Community.Display.V1.FormattedValue"] || "").trim().toLowerCase();

            return filterCriteria.assignTo.some(assignee =>

                assignedTo === assignee.trim().toLowerCase()

            );

        });

    }



    // Apply approval status filter

    if (filterCriteria.approvalStatus && filterCriteria.approvalStatus.length > 0) {

        filteredData = filteredData.filter(result => {

            const requestStatus = (result["al_requeststatus@OData.Community.Display.V1.FormattedValue"] || "").trim().toLowerCase();

            return filterCriteria.approvalStatus.some(status =>

                requestStatus === status.trim().toLowerCase()

            );

        });

    }



    // Apply final status filter

    if (filterCriteria.finalStatus && filterCriteria.finalStatus.length > 0) {

        filteredData = filteredData.filter(result => {

            const finalStatus = (result["al_finalrequeststatus@OData.Community.Display.V1.FormattedValue"] || "").trim().toLowerCase();

            return filterCriteria.finalStatus.some(status =>

                finalStatus === status.trim().toLowerCase()

            );

        });

    }



    // Apply date filter

    if (filterCriteria.minDate || filterCriteria.maxDate) {

        filteredData = filteredData.filter(result => {

            if (!result["createdon"]) return false;

            const requestDate = moment(result["createdon"]);

            if (filterCriteria.minDate && requestDate.isBefore(filterCriteria.minDate, 'day')) return false;

            if (filterCriteria.maxDate && requestDate.isAfter(filterCriteria.maxDate, 'day')) return false;

            return true;

        });

    }



    // Store filtered data

    filteredRequestData = filteredData;



    // Clear table and rebind with filtered data (first 50 items)

    tableElement.clear();

    const firstBatch = filteredData.slice(0, RECORDS_PER_BATCH);

    await BindRequest(userWithTeam, firstBatch, tableElement, teamType, null, null, false, true);



    // Update count

    //$(".pendingReq").text(filteredData.length);



    // Reset batch counter

    loadedPendingBatches = 1;



    // Update allPendingRequests to filtered data for batch loading

    allPendingRequests = filteredData;



    // Update pagination to reflect filtered count - update immediately after table is updated
    if (tableElement) {
        // Update pagination settings immediately
        window.updateDataTablesPagination(tableElement, filteredData.length);

        // Also update the info text manually to ensure it's visible
        setTimeout(() => {
            try {
                if (typeof tableElement.api === 'function') {
                    let api = tableElement.api();
                    if (api) {
                        let pageInfo = api.page.info();
                        let start = pageInfo.recordsDisplay == 0 ? 0 : pageInfo.start + 1;
                        let end = pageInfo.end;
                        let total = filteredData.length;
                        let $info = $(api.table().node()).parents('.table-responsive').find('.dt-info');
                        if ($info.length) {
                            $info.text(`Showing ${start} to ${end} of ${total} entries`);
                        }
                        // Force a redraw to update pagination numbers
                        api.draw(false);
                    }
                }
            } catch (error) {
                console.warn('Error updating pagination info:', error);
            }
        }, 150);
    }



    // Update KPI cards with filtered data counts

    // if (filteredData && filteredData.length > 0) {

    //     getRequestCount(filteredData);
    // }
}

// Debounce timer for search input

let searchInputTimeout = null;
$(document).on("input keyup", "#searchText", async function () {

    clearTimeout(searchInputTimeout);



    // Debounce search to avoid too many calls (300ms delay)

    searchInputTimeout = setTimeout(async () => {

        let searchText = $(this).val()?.trim() || null;



        // If we have stored data, use filter function

        if (allRequestDataStorage && allRequestDataStorage.length > 0) {

            var table;

            var userWithTeam;

            var teamType = -1;



            // Get current user and team info

            try {

                const loggedInUserId = $("#userIdHidden").html();

                if (loggedInUserId) {

                    userWithTeam = await GetUserWithTeam(loggedInUserId);

                    if (userWithTeam) {

                        teamType = userWithTeam.al_teamtype;

                    }

                }

            } catch (error) {

                console.error("Error getting user with team:", error);

            }



            if ($('#pendingReqTable').is(":visible")) {

                table = $('#pendingReqTable').DataTable();

            } else if ($('#grcApproveTable').is(":visible")) {

                table = $('#grcApproveTable').DataTable();

            }

            else if ($('#grcRejectTable').is(":visible")) {

                table = $('#grcRejectTable').DataTable();

            }

            else if ($('#returenedRequestTable').is(":visible")) {

                table = $('#returenedRequestTable').DataTable();

            }

            else if ($('#p1Table').is(":visible")) {

                table = $('#p1Table').DataTable();

            }

            else {

                table = $('#allReqTable').DataTable();

            }



            // Get other filter values

            var requestSubType = $('#requestSubtype').val() || [];

            var legalEntity = $("#legalEntityFilter").val() || [];

            var assignTo = $("#assignToFilter").val() || [];

            var approvalStatus = $('#approvalStatus').val() || [];

            var finalStatus = $("#finalStatus").val() || [];



            // Date filter

            var dateFilter = $('#dateFilter').val();

            var fromDate = $('#filterFromDate').val();

            var toDate = $('#filterToDate').val();

            let minDate = null, maxDate = null;

            if (dateFilter === 'last7') {

                minDate = moment().subtract(7, 'days').format('YYYY-MM-DD');

            } else if (dateFilter === 'last15') {

                minDate = moment().subtract(15, 'days').format('YYYY-MM-DD');

            } else if (dateFilter === 'last30') {

                minDate = moment().subtract(30, 'days').format('YYYY-MM-DD');

            } else if (dateFilter === 'custom') {

                minDate = fromDate || null;

                maxDate = toDate || null;

            }



            // If search is empty and no other filters, reset to original data

            if (!searchText && !requestSubType.length && !legalEntity.length && !assignTo.length &&

                !approvalStatus.length && !finalStatus.length && !minDate && !maxDate) {

                //allPendingRequests = allRequestDataStorage;

                filteredRequestData = null;

                table.clear();

                const firstBatch = allRequestDataStorage.slice(0, RECORDS_PER_BATCH);

                await BindRequest(userWithTeam, firstBatch, table, -1, null, null, false, true);

                // $(".pendingReq").text(allRequestDataStorage.length);

                loadedPendingBatches = 1;



                // Update KPI cards with original data counts

                // if (allRequestDataStorage && allRequestDataStorage.length > 0) {

                //     getRequestCount(allRequestDataStorage);

                // }

                return;

            }



            // Apply filters including search text

            await applyFiltersToStoredData({

                searchText: searchText,

                requestSubType: requestSubType,

                legalEntity: legalEntity,

                assignTo: assignTo,

                approvalStatus: approvalStatus,

                finalStatus: finalStatus,

                minDate: minDate ? moment(minDate) : null,

                maxDate: maxDate ? moment(maxDate) : null

            }, userWithTeam, table, -1);



            return;

        }



        // Fallback to original DataTables search for other cases

        // clear previous search filters

        $.fn.dataTable.ext.search = [];


        $.fn.dataTable.ext.search.push(function (settings, data) {

            let matchSearchText;


            if ($('#pendingReqTable').is(":visible")) {

                matchSearchText = !searchText ||

                    data[1].toLowerCase().includes(searchText) ||  // Request ID

                    data[2].toLowerCase().includes(searchText) ||  // UVP Request ID

                    data[3].toLowerCase().includes(searchText) ||  // Request Sub Type

                    data[4].toLowerCase().includes(searchText) ||  // Req Raised Date

                    data[5].toLowerCase().includes(searchText) ||  // Vendor Name

                    data[6].toLowerCase().includes(searchText) ||  // Legal Entity

                    data[7].toLowerCase().includes(searchText) ||  // Assign To

                    data[8].toLowerCase().includes(searchText) ||  // RequestedBy

                    data[9].toLowerCase().includes(searchText) ||  // Assigned Date & Time

                    data[10].toLowerCase().includes(searchText) || // Due Date

                    data[11].toLowerCase().includes(searchText) || // Approval Status

                    data[12].toLowerCase().includes(searchText) || // Final Request Status

                    data[14].toLowerCase().includes(searchText) || // PR Number

                    data[15].toLowerCase().includes(searchText) || // PO Number

                    data[16].toLowerCase().includes(searchText) || // ProjectID

                    data[18].toLowerCase().includes(searchText);   // EAuctionName

            } else {

                matchSearchText = !searchText ||

                    data[0].toLowerCase().includes(searchText) ||  // Request ID

                    data[1].toLowerCase().includes(searchText) ||  // UVP Request ID

                    data[2].toLowerCase().includes(searchText) ||  // Request Sub Type

                    data[3].toLowerCase().includes(searchText) ||  // Req Raised Date

                    data[4].toLowerCase().includes(searchText) ||  // Vendor Name

                    data[5].toLowerCase().includes(searchText) ||  // Legal Entity

                    data[6].toLowerCase().includes(searchText) ||  // Assign To

                    data[7].toLowerCase().includes(searchText) ||  // RequestedBy

                    data[8].toLowerCase().includes(searchText) ||  // Assigned Date & Time

                    data[9].toLowerCase().includes(searchText) ||  // Due Date

                    data[10].toLowerCase().includes(searchText) || // Approval Status

                    data[11].toLowerCase().includes(searchText) || // Final Request Status

                    data[13].toLowerCase().includes(searchText) || // PR Number

                    data[14].toLowerCase().includes(searchText) || // PO Number

                    data[15].toLowerCase().includes(searchText) || // ProjectID

                    data[17].toLowerCase().includes(searchText);   // EAuctionName

            }


            return matchSearchText;

        });


        // redraw the currently visible table

        if ($('#pendingReqTable').is(":visible")) $('#pendingReqTable').DataTable().draw();

        else if ($('#grcApproveTable').is(":visible")) $('#grcApproveTable').DataTable().draw();

        else if ($('#grcRejectTable').is(":visible")) $('#grcRejectTable').DataTable().draw();

        else if ($('#returenedRequestTable').is(":visible")) $('#returenedRequestTable').DataTable().draw();

        else if ($('#p1Table').is(":visible")) $('#p1Table').DataTable().draw();

        else $('#allReqTable').DataTable().draw();

    }, 300); // 300ms debounce

});


$(document).on('change', '#dateFilter', function () {

    const selected = $(this).val();

    if (selected === 'custom') {

        $("#customDateFilterModal").modal('show');

    } else {

        $("#customDateFilterModal").modal('hide');

        $('#filterFromDate').val('');

        $('#filterToDate').val('');

    }
});


$(document).on("click", "#resetBtn", async function () {

    // Show loader for better UX

    $('.custom-loader').css("display", "flex");



    try {

        // Clear all filter fields

        // Note: We don't trigger 'change' events as filtering only works on button click and search keyup

        $('#requestSubtype').val(null).trigger("change");



        $('#legalEntityFilter').val(null).trigger("change");



        $('#assignToFilter').val(null).trigger("change");



        $('#approvalStatus').val(null).trigger("change");



        $('#finalStatus').val(null).trigger("change");



        $('#searchText').val('');



        $('#dateFilter').val('').trigger("change");







        $('#filterFromDate').val('');



        $('#filterToDate').val('');







        // Clear any DataTables search filters



        $.fn.dataTable.ext.search = [];







        // If we have stored data, reset to show all original data



        if (allRequestDataStorage && allRequestDataStorage.length > 0) {



            var table;



            var userWithTeam;



            var teamType = -1;







            // Get current user and team info

            // Ensure global loggedInUserId is set for getRequestCount function

            const currentUserIdForTeam = $("#userIdHidden").html();

            if (currentUserIdForTeam) {

                loggedInUserId = String(currentUserIdForTeam).trim(); // Set global variable for getRequestCount
            }
            try {
                if (currentUserIdForTeam) {
                    userWithTeam = await GetUserWithTeam(currentUserIdForTeam);
                    if (userWithTeam) {
                        teamType = userWithTeam.al_teamtype;
                    }
                }
            } catch (error) {
                console.error("Error getting user with team:", error);
            }
            if ($('#pendingReqTable').is(":visible")) {
                table = $('#pendingReqTable').DataTable();
            } else if ($('#grcApproveTable').is(":visible")) {
                table = $('#grcApproveTable').DataTable();
            }
            else if ($('#grcRejectTable').is(":visible")) {
                table = $('#grcRejectTable').DataTable();
            }
            else if ($('#returenedRequestTable').is(":visible")) {
                table = $('#returenedRequestTable').DataTable();
            }
            else if ($('#p1Table').is(":visible")) {
                table = $('#p1Table').DataTable();
            }
            else {
                table = $('#allReqTable').DataTable();
            }
            // Reset to original data
            // allPendingRequests = allRequestDataStorage;


            filteredRequestData = null;



            // Clear the table and reload with original data

            table.clear();


            const firstBatch = allRequestDataStorage.slice(0, RECORDS_PER_BATCH);


            await BindRequest(userWithTeam, firstBatch, table, -1, null, null, false, true);



            // Reset ALL batch loading variables for all tabs
            loadedPendingBatches = 1;
            loadedAllBatches = 1;
            loadedResolvedBatches = 1;
            loadedReturnedBatches = 1;
            loadedRejectedBatches = 1;
            loadedP1Batches = 1;
            
            // Reset loading flags
            isLoadingBatch = false;
            isLoadingAllBatch = false;
            isLoadingResolvedBatch = false;
            isLoadingReturnedBatch = false;
            isLoadingRejectedBatch = false;
            isLoadingP1Batch = false;
            
            // Reset last checked pages
            lastCheckedPage = -1;
            lastCheckedAllPage = -1;
            lastCheckedResolvedPage = -1;
            lastCheckedReturnedPage = -1;
            lastCheckedRejectedPage = -1;
            lastCheckedP1Page = -1;
            
            // Clear any pending batch load timeouts
            if (batchLoadTimeout) clearTimeout(batchLoadTimeout);
            if (allAllBatchLoadTimeout) clearTimeout(allAllBatchLoadTimeout);
            if (allResolvedBatchLoadTimeout) clearTimeout(allResolvedBatchLoadTimeout);
            if (allReturnedBatchLoadTimeout) clearTimeout(allReturnedBatchLoadTimeout);
            if (allRejectedBatchLoadTimeout) clearTimeout(allRejectedBatchLoadTimeout);
            if (allP1BatchLoadTimeout) clearTimeout(allP1BatchLoadTimeout);
            
            // Note: KPI cards should NOT be updated here - they should always show total counts

            // KPI counts are set when data is initially loaded, not when clearing filters

            // The KPI cards should remain unchanged and display counts from allRequestsForKPICounts
            // Update pagination to reflect original count
            setTimeout(() => {
                if (table && allRequestDataStorage.length > 0) {
                    window.updateDataTablesPagination(table, allRequestDataStorage.length);
                }
            }, 100);


            // Redraw table to ensure it's updated

            table.draw();



            // Force Select2 dropdowns to update their display (only for dateFilter which uses Select2)

            if ($('#dateFilter').hasClass('select2-hidden-accessible')) {

                $('#dateFilter').trigger('change.select2');

            }




            // Hide loader after operation completes

            $('.custom-loader').css("display", "none");





            return;



        }



        // Original reset logic for other cases - redraw all tables





        if ($('#pendingReqTable').is(":visible")) {



            $('#pendingReqTable').DataTable().draw();



        } else if ($('#grcApproveTable').is(":visible")) {



            $('#grcApproveTable').DataTable().draw();



        }



        else if ($('#grcRejectTable').is(":visible")) {



            $('#grcRejectTable').DataTable().draw();



        }



        else if ($('#returenedRequestTable').is(":visible")) {



            $('#returenedRequestTable').DataTable().draw();



        }



        else if ($('#p1Table').is(":visible")) {



            $('#p1Table').DataTable().draw();



        }



        else {
            $('#allReqTable').DataTable().draw();
        }
        // Hide loader after all operations complete
        $('.custom-loader').css("display", "none");
    } catch (error) {

        console.error("Error in clear button:", error);

        $('.custom-loader').css("display", "none");
    }
});

// $(document).on("click", "#resetBtn", async function () {
//     // Clear all filter fields

//     $('#requestSubtype').val(null).trigger('change');

//     $('#legalEntityFilter').val(null).trigger('change');

//     $('#assignToFilter').val(null).trigger('change');

//     $('#approvalStatus').val(null).trigger('change');

//     $('#finalStatus').val(null).trigger('change');

//     $('#searchText').val('');

//     $('#dateFilter').val('').trigger('change');

//     $('#filterFromDate').val('');

//     $('#filterToDate').val('');



//     // Clear any DataTables search filters

//     $.fn.dataTable.ext.search = [];



//     // If we have stored data, reset to show all original data

//     if (allRequestDataStorage && allRequestDataStorage.length > 0) {

//         var table;

//         var userWithTeam;

//         var teamType = -1;



//         // Get current user and team info

//         try {

//             const loggedInUserId = $("#userIdHidden").html();

//             if (loggedInUserId) {

//                 userWithTeam = await GetUserWithTeam(loggedInUserId);

//                 if (userWithTeam) {

//                     teamType = userWithTeam.al_teamtype;

//                 }

//             }

//         } catch (error) {

//             console.error("Error getting user with team:", error);

//         }



//         if ($('#pendingReqTable').is(":visible")) {

//             table = $('#pendingReqTable').DataTable();

//         } else if ($('#grcApproveTable').is(":visible")) {

//             table = $('#grcApproveTable').DataTable();

//         }

//         else if ($('#grcRejectTable').is(":visible")) {

//             table = $('#grcRejectTable').DataTable();

//         }

//         else if ($('#returenedRequestTable').is(":visible")) {

//             table = $('#returenedRequestTable').DataTable();

//         }

//         else if ($('#p1Table').is(":visible")) {

//             table = $('#p1Table').DataTable();

//         }

//         else {

//             table = $('#allReqTable').DataTable();

//         }



//         // Reset to original data

//         allPendingRequests = allRequestDataStorage;

//         filteredRequestData = null;

//         table.clear();

//         const firstBatch = allRequestDataStorage.slice(0, RECORDS_PER_BATCH);

//         await BindRequest(userWithTeam, firstBatch, table, -1, null, null, false, true);

//         $(".pendingReq").text(allRequestDataStorage.length);

//         loadedPendingBatches = 1;

//         // Update pagination to reflect original count
//         setTimeout(() => {
//             if (table && allRequestDataStorage.length > 0) {
//                 window.updateDataTablesPagination(table, allRequestDataStorage.length);
//             }
//         }, 100);

//         // Update KPI cards with original data counts

//         if (allRequestDataStorage && allRequestDataStorage.length > 0) {

//             getRequestCount(allRequestDataStorage);

//         }



//         // Redraw table to ensure it's updated

//         table.draw();

//         return;

//     }



//     // Original reset logic for other cases - redraw all tables

//     if ($('#pendingReqTable').is(":visible")) {

//         $('#pendingReqTable').DataTable().draw();

//     } else if ($('#grcApproveTable').is(":visible")) {

//         $('#grcApproveTable').DataTable().draw();

//     }

//     else if ($('#grcRejectTable').is(":visible")) {

//         $('#grcRejectTable').DataTable().draw();

//     }

//     else if ($('#returenedRequestTable').is(":visible")) {

//         $('#returenedRequestTable').DataTable().draw();

//     }

//     else if ($('#p1Table').is(":visible")) {

//         $('#p1Table').DataTable().draw();

//     }

//     else {
//         $('#allReqTable').DataTable().draw();
//     }
// });

// Filter button click handler - calls the common function
$(document).on("click", "#filterBtn", async function () {
    await applyFiltersAndUpdatePagination();
});

//Nottification Functionality
async function getNotificationItems() {
    await $.ajax({
        type: "GET",
        url: getAllNotificationUrl + '&teamType=' + currentUserTeamType + '&userId=' + contactIdGlobal,
        success: function (response) {
            if (response.length > 0) {
                let allNotifications = [];

                // Step 1: Flatten and collect all messages with requestId
                response.forEach(req => {
                    req.requestNotifications.forEach(note => {
                        allNotifications.push({
                            msg: note.chatMsg,
                            requestId: req.requestId,
                            messageId: note.id,
                            createdOn: new Date(note.createdOn)
                        });
                    });
                });

                // Step 2: Sort by createdOn descending
                allNotifications.sort((a, b) => b.createdOn - a.createdOn);

                // Step 3: Take top 3
                const topNotifications = allNotifications.slice(0, 3);

                // Step 4: Render
                $('#notification-container').empty();
                let reqNotificationIdToMarkRead = [];
                topNotifications.forEach(note => {
                    $('#notification-container').append(`
                    <div class="col-sm-12 mb-4">
                        <div class="d-flex align-items-center justify-content-between card_box_main">
                            <p class="m-0 notification_message">⏳Reminder:${note.msg}</p>
                            <a href="/request-journey/?id=${note.requestId}&view=true" class="view_notification">View</a>
                        </div>
                    </div>
                `);
                    reqNotificationIdToMarkRead.push(note.messageId);
                });
                // Step 5: Mark these messages as read
                if (reqNotificationIdToMarkRead.length > 0) {
                    let readPayload = {
                        TeamType: 1,
                        RequestChatIds: reqNotificationIdToMarkRead,
                    };

                    $.ajax({
                        type: "POST",
                        url: markAsReadNotificaitonUrl,
                        data: JSON.stringify(readPayload),
                        contentType: "application/json",
                        success: function (response) {
                            console.log("Marked as read:", response);
                        },
                        error: function (error) {
                            console.log("Error marking as read:", error);
                            showErrorModal("Error marking as read");
                        }
                    });
                }
            }
        },
        error: function (error) {
            console.log(error);
            showErrorModal("Get Notification :- Notification not get");
        }
    });
}

// Custom method to validate date range
$.validator.addMethod("endDateAfterStart", function (value, element) {
    const startDate = $('#startDate').val();
    if (!startDate || !value) return true; // Let "required" handle empty fields
    return new Date(value) >= new Date(startDate);
}, "End date must be after or equal to start date.");
// Initialize form validation
$('#customDateForm').validate({
    rules: {
        startDate: {
            required: true,
            date: true
        },
        endDate: {
            required: true,
            date: true,
            endDateAfterStart: true // Custom rule
        }
    },
    messages: {
        startDate: {
            required: "Please enter a start date.",
            date: "Please enter a valid date."
        },
        endDate: {
            required: "Please enter an end date.",
            date: "Please enter a valid date.",
            endDateAfterStart: "End date must be after or equal to start date."
        }
    },
    errorPlacement: function (error, element) {
        error.addClass('invalid-feedback');
        if (element.parent('.input-group').length) {
            error.insertAfter(element.parent());
        } else {
            error.insertAfter(element);
        }
    },
});

// Custom method to validate date range
$.validator.addMethod("endDateAfterStart", function (value, element) {
    const startDate = $('#filterFromDate').val();
    if (!startDate || !value) return true; // Let "required" handle empty fields
    return new Date(value) >= new Date(startDate);
}, "End date must be after or equal to start date.");

// Initialize form validation
$('#customDateFilterForm').validate({
    rules: {
        filterFromDate: {
            required: true,
            date: true
        },
        filterToDate: {
            required: true,
            date: true,
            endDateAfterStart: true // Custom rule
        }
    },
    messages: {
        filterFromDate: {
            required: "Please enter a start date.",
            date: "Please enter a valid date."
        },
        filterToDate: {
            required: "Please enter an end date.",
            date: "Please enter a valid date.",
            endDateAfterStart: "End date must be after or equal to start date."
        }
    },
    errorPlacement: function (error, element) {
        error.addClass('invalid-feedback');
        if (element.parent('.input-group').length) {
            error.insertAfter(element.parent());
        } else {
            error.insertAfter(element);
        }
    },
});

$('#startDate, #endDate, #filterFromDate, #filterToDate').on('change', function () {
    $(this).valid(); // Re-run validation just for this field
});

function getTableByTabName(tabName) {
    if (!tabName) return pendingReqTable; // Default to pending table

    const tabNameLower = tabName;

    // Map tab names to their corresponding table variables
    if (tabNameLower.includes('All') || tabNameLower === 'all requests') {
        return allReqTable;
    } else if (tabNameLower.includes('pending') || tabNameLower === 'pending requests') {
        return pendingReqTable;
    } else if (tabNameLower.includes('Returned') || tabNameLower === 'returned requests') {
        return returenedRequestTable;
    } else if (tabNameLower.includes('approved') || tabNameLower.includes('Resolved') || tabNameLower === 'approved requests') {
        return grcApproveTable;
    } else if (tabNameLower.includes('Rejected') || tabNameLower === 'rejected requests') {
        return grcRejectTable;
    }

    // Default to pending table if no match
    return pendingReqTable;
}

//updatee9090