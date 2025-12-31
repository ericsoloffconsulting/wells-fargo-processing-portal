/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 */
define(['N/ui/serverWidget', 'N/search', 'N/log', 'N/url', 'N/record', 'N/redirect', 'N/query', 'N/email', 'N/runtime'], function (serverWidget, search, log, url, record, redirect, query, email, runtime) {

    /**
  * Handles GET and POST requests to the Suitelet
  * @param {Object} context - NetSuite context object containing request/response
  * @returns {void}
  */
    function onRequest(context) {
        if (context.request.method === 'GET') {
            var form = serverWidget.createForm({
                title: 'Wells Fargo Processing'
            });

            try {
                // Pass context to buildSearchResultsHTML
                var htmlContent = buildSearchResultsHTML(context);

                // Add the HTML field to display the search results
                var htmlField = form.addField({
                    id: 'custpage_search_results',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: 'Search Results'
                });
                htmlField.defaultValue = htmlContent;

                // Add a refresh button
                form.addButton({
                    id: 'custpage_refresh',
                    label: 'Refresh',
                    functionName: 'refreshPage'
                });

            } catch (e) {
                log.error('Error in Wells Fargo Processing Suitelet', e.message);
                var errorField = form.addField({
                    id: 'custpage_error',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: 'Error'
                });
                errorField.defaultValue = '<div style="color: red;">Error loading search results: ' + e.message + '</div>';
            }

            context.response.writePage(form);

        } else if (context.request.method === 'POST') {
            try {
                var action = context.request.parameters.action;

                if (action === 'create_deposit') {
                    var customerId = context.request.parameters.customer;
                    var amount = context.request.parameters.amount;
                    var wfAuthId = context.request.parameters.wfAuthId;
                    var salesOrderId = context.request.parameters.salesorder;
                    var departmentId = context.request.parameters.location;
                    var wfAuthNumber = context.request.parameters.wfAuthNumber;

                    log.debug('Creating Customer Deposit - Input Values', {
                        customer: customerId,
                        amount: amount,
                        wfAuthId: wfAuthId,
                        salesOrder: salesOrderId,
                        departmentId: departmentId,
                        wfAuthNumber: wfAuthNumber
                    });

                    // Validate inputs
                    if (!customerId || !amount) {
                        throw new Error('Missing required parameters: customer=' + customerId + ', amount=' + amount);
                    }

                    // Parse and validate numeric values
                    var customerIdInt = parseInt(customerId, 10);
                    var amountFloat = parseFloat(amount);
                    var salesOrderIdInt = salesOrderId ? parseInt(salesOrderId, 10) : null;
                    var departmentIdInt = departmentId ? parseInt(departmentId, 10) : 1;

                    if (isNaN(customerIdInt) || customerIdInt <= 0) {
                        throw new Error('Invalid customer ID: ' + customerId);
                    }

                    if (isNaN(amountFloat) || amountFloat <= 0) {
                        throw new Error('Invalid amount: ' + amount);
                    }

                    // Lookup fulfilling location from department record
                    var fulfillingLocationId = 1;
                    if (departmentIdInt) {
                        try {
                            fulfillingLocationId = lookupFulfillingLocation(departmentIdInt);
                        } catch (lookupError) {
                            log.error('Error looking up fulfilling location', {
                                error: lookupError.message,
                                departmentId: departmentIdInt
                            });
                        }
                    }

                    log.debug('Validated input values', {
                        customerIdInt: customerIdInt,
                        amountFloat: amountFloat,
                        salesOrderIdInt: salesOrderIdInt,
                        departmentIdInt: departmentIdInt,
                        fulfillingLocationId: fulfillingLocationId
                    });

                    // Create the Customer Deposit record
                    var depositRecord = record.create({
                        type: record.Type.CUSTOMER_DEPOSIT,
                        isDynamic: true
                    });

                    try {
                        // Set customer field
                        depositRecord.setValue({
                            fieldId: 'customer',
                            value: customerIdInt
                        });

                        // Set location (fulfilling location)
                        depositRecord.setValue({
                            fieldId: 'location',
                            value: fulfillingLocationId
                        });

                        // Set department (selling location)
                        depositRecord.setValue({
                            fieldId: 'department',
                            value: departmentIdInt
                        });

                        // Set transaction date
                        depositRecord.setValue({
                            fieldId: 'trandate',
                            value: new Date()
                        });

                        // Enhanced memo with more details
                        var memoText = 'Wells Fargo Customer Deposit - WF Auth #: ' + wfAuthNumber;
                        memoText += ' - WF Record ID: WF' + wfAuthId;

                        depositRecord.setValue({
                            fieldId: 'memo',
                            value: memoText
                        });

                        // Set the Wells Fargo Authorization link in custom body field
                        if (wfAuthId) {
                            depositRecord.setValue({
                                fieldId: 'custbody_linked_wells_fargo_authorizat',
                                value: parseInt(wfAuthId, 10)
                            });
                        }

                        // Set Sales Order reference if available
                        if (salesOrderIdInt) {
                            depositRecord.setValue({
                                fieldId: 'salesorder',
                                value: salesOrderIdInt
                            });
                        }

                        // Set payment amount
                        depositRecord.setValue({
                            fieldId: 'payment',
                            value: amountFloat
                        });

                        // Set payment method (ACH)
                        depositRecord.setValue({
                            fieldId: 'paymentmethod',
                            value: 12
                        });

                        // Save the Customer Deposit
                        var depositId = depositRecord.save();

                        log.audit('Customer Deposit Created Successfully', {
                            depositId: depositId,
                            customer: customerIdInt,
                            amount: amountFloat,
                            location: fulfillingLocationId,
                            department: departmentIdInt,
                            salesOrder: salesOrderIdInt,
                            wfAuthId: wfAuthId,
                            wfAuthNumber: wfAuthNumber
                        });

                        // Get the transaction ID from the newly created deposit
                        var depositTranId = '';
                        var wfAuthName = '';

                        try {
                            var savedDepositRecord = record.load({
                                type: record.Type.CUSTOMER_DEPOSIT,
                                id: depositId,
                                isDynamic: false
                            });
                            depositTranId = savedDepositRecord.getValue('tranid') || depositId;
                        } catch (loadError) {
                            log.error('Error loading deposit record for tranid', loadError.message);
                            depositTranId = depositId; // Fallback to ID
                        }

                        // Update the Wells Fargo Authorization record and get its name
                        if (wfAuthId) {
                            try {
                                // Get existing deposit links
                                var existingDepositLinks = getExistingDepositLinks(wfAuthId);
                                var updatedDepositLinks = appendDepositToMultipleSelect(existingDepositLinks, depositId);

                                // Load Wells Fargo Authorization record to get current amounts
                                var wfAuthRecord = record.load({
                                    type: 'customrecord_bas_wf_auth',
                                    id: wfAuthId,
                                    isDynamic: false
                                });
                                wfAuthName = wfAuthRecord.getValue('name') || wfAuthId;

                                // Get current "To Be Charged" amount (treat null as 0)
                                var currentToBeCharged = wfAuthRecord.getValue('custrecord_wf_deposit_to_be_charged') || 0;
                                currentToBeCharged = parseFloat(currentToBeCharged);

                                // Get current "Charged" amount (treat null as 0)
                                var currentCharged = wfAuthRecord.getValue('custrecord_wells_fargo_auth_dep_charged') || 0;
                                currentCharged = parseFloat(currentCharged);

                                // Calculate new amounts
                                var newToBeCharged = currentToBeCharged - amountFloat;
                                var newCharged = currentCharged + amountFloat;

                                log.debug('Updating Wells Fargo Auth amounts', {
                                    wfAuthId: wfAuthId,
                                    depositAmount: amountFloat,
                                    currentToBeCharged: currentToBeCharged,
                                    newToBeCharged: newToBeCharged,
                                    currentCharged: currentCharged,
                                    newCharged: newCharged
                                });

                                // Update the Wells Fargo Authorization record with all changes
                                record.submitFields({
                                    type: 'customrecord_bas_wf_auth',
                                    id: wfAuthId,
                                    values: {
                                        'custrecord_customer_deposit_link': updatedDepositLinks,
                                        'custrecord_wf_deposit_to_be_charged': newToBeCharged,
                                        'custrecord_wells_fargo_auth_dep_charged': newCharged
                                    }
                                });

                                log.audit('Wells Fargo Auth record updated successfully', {
                                    wfAuthId: wfAuthId,
                                    wfAuthName: wfAuthName,
                                    depositId: depositId,
                                    depositTranId: depositTranId,
                                    updatedDepositLinks: updatedDepositLinks,
                                    newToBeCharged: newToBeCharged,
                                    newCharged: newCharged
                                });

                            } catch (updateError) {
                                log.error('Error updating Wells Fargo Auth record', {
                                    error: updateError.message,
                                    stack: updateError.stack,
                                    wfAuthId: wfAuthId,
                                    depositId: depositId
                                });
                                // Set fallback values for success message
                                wfAuthName = wfAuthId;
                            }
                        }

                        // Redirect back to the same page with success message
                        redirect.toSuitelet({
                            scriptId: context.request.parameters.script,
                            deploymentId: context.request.parameters.deploy,
                            parameters: {
                                success: 'true',
                                depositTranId: depositTranId,
                                wfAuthName: wfAuthName
                            }
                        });

                    } catch (fieldError) {
                        log.error('Error setting fields on Customer Deposit', {
                            error: fieldError.message,
                            stack: fieldError.stack,
                            customer: customerIdInt,
                            amount: amountFloat,
                            departmentId: departmentIdInt,
                            fulfillingLocationId: fulfillingLocationId
                        });
                        throw fieldError;
                    }

                } else if (action === 'create_payment') {
                    var customerId = context.request.parameters.customer;
                    var amount = context.request.parameters.amount;
                    var wfAuthNumber = context.request.parameters.wfAuthNumber;
                    var invoiceNumber = context.request.parameters.invoiceNumber;

                    log.debug('Creating Customer Payment - Input Values', {
                        customer: customerId,
                        amount: amount,
                        wfAuthNumber: wfAuthNumber,
                        invoiceNumber: invoiceNumber
                    });

                    // Validate inputs
                    if (!customerId || !amount) {
                        throw new Error('Missing required parameters: customer=' + customerId + ', amount=' + amount);
                    }

                    // Parse and validate numeric values
                    var customerIdInt = parseInt(customerId, 10);
                    var amountFloat = parseFloat(amount);

                    if (isNaN(customerIdInt) || customerIdInt <= 0) {
                        throw new Error('Invalid customer ID: ' + customerId);
                    }

                    if (isNaN(amountFloat) || amountFloat <= 0) {
                        throw new Error('Invalid amount: ' + amount);
                    }

                    // Find the invoice to apply payment to
                    var invoiceId = null;
                    if (invoiceNumber) {
                        try {
                            invoiceId = findInvoiceByNumber(customerIdInt, invoiceNumber);
                            log.debug('Found invoice', {
                                invoiceNumber: invoiceNumber,
                                invoiceId: invoiceId
                            });
                        } catch (findError) {
                            log.error('Error finding invoice', {
                                error: findError.message,
                                invoiceNumber: invoiceNumber,
                                customerId: customerIdInt
                            });
                        }
                    }

                    // Create Customer Payment using transform if we have an invoice, otherwise create new
                    var paymentRecord;
                    var isTransformed = false;

                    if (invoiceId) {
                        try {
                            // Transform the invoice into a customer payment
                            paymentRecord = record.transform({
                                fromType: record.Type.INVOICE,
                                fromId: invoiceId,
                                toType: record.Type.CUSTOMER_PAYMENT,
                                isDynamic: true  // Required for dynamic sublist manipulation
                            });
                            isTransformed = true;

                            log.debug('Transformed invoice to payment', {
                                invoiceId: invoiceId,
                                invoiceNumber: invoiceNumber
                            });

                        } catch (transformError) {
                            log.error('Error transforming invoice to payment', {
                                error: transformError.message,
                                invoiceId: invoiceId
                            });

                            // Fallback to creating new payment
                            paymentRecord = record.create({
                                type: record.Type.CUSTOMER_PAYMENT,
                                isDynamic: true
                            });
                            isTransformed = false;
                        }
                    } else {
                        // Create new payment if no invoice found
                        paymentRecord = record.create({
                            type: record.Type.CUSTOMER_PAYMENT,
                            isDynamic: true
                        });
                        isTransformed = false;
                    }

                    try {
                        // Set customer field (if not already set by transform)
                        if (!isTransformed) {
                            paymentRecord.setValue({
                                fieldId: 'customer',
                                value: customerIdInt
                            });
                        }

                        // Set transaction date
                        paymentRecord.setValue({
                            fieldId: 'trandate',
                            value: new Date()
                        });

                        // Set payment method (ACH)
                        paymentRecord.setValue({
                            fieldId: 'paymentmethod',
                            value: 12
                        });

                        // Set memo with Wells Fargo information
                        var memoText = 'Wells Fargo Payment - Auth # ' + (wfAuthNumber || 'Unknown');
                        paymentRecord.setValue({
                            fieldId: 'memo',
                            value: memoText
                        });

                        // CRITICAL: Set payment amount BEFORE applying to invoice
                        // This ensures the payment header has the correct amount available
                        paymentRecord.setValue({
                            fieldId: 'payment',
                            value: amountFloat
                        });

                        log.debug('Payment amount set', {
                            amount: amountFloat,
                            isTransformed: isTransformed
                        });

                        // Apply payment to specific invoice if found and transformed
                        if (invoiceId && isTransformed) {
                            try {
                                // Get apply sublist line count
                                var applyLineCount = paymentRecord.getLineCount({
                                    sublistId: 'apply'
                                });

                                log.debug('Apply sublist info', {
                                    lineCount: applyLineCount,
                                    targetInvoiceId: invoiceId,
                                    isTransformed: isTransformed,
                                    paymentAmount: amountFloat
                                });

                                // For transformed records, the source invoice should already be on line 0
                                // We need to update the amount being applied using DYNAMIC MODE methods
                                var foundInvoice = false;

                                for (var line = 0; line < applyLineCount; line++) {
                                    // Select the line to work with it (DYNAMIC MODE REQUIRED)
                                    paymentRecord.selectLine({
                                        sublistId: 'apply',
                                        line: line
                                    });

                                    // Get the internal ID of the document on this line
                                    var applyInternalId = paymentRecord.getCurrentSublistValue({
                                        sublistId: 'apply',
                                        fieldId: 'doc'
                                    });

                                    log.debug('Checking apply line', {
                                        line: line,
                                        applyInternalId: applyInternalId,
                                        targetInvoiceId: invoiceId
                                    });

                                    if (parseInt(applyInternalId, 10) === parseInt(invoiceId, 10)) {
                                        foundInvoice = true;

                                        // Get the current apply status
                                        var isCurrentlyApplied = paymentRecord.getCurrentSublistValue({
                                            sublistId: 'apply',
                                            fieldId: 'apply'
                                        });

                                        log.debug('Found target invoice on apply sublist', {
                                            line: line,
                                            invoiceId: invoiceId,
                                            isCurrentlyApplied: isCurrentlyApplied,
                                            requestedAmount: amountFloat
                                        });

                                        // Set apply flag to true (if not already)
                                        if (!isCurrentlyApplied) {
                                            paymentRecord.setCurrentSublistValue({
                                                sublistId: 'apply',
                                                fieldId: 'apply',
                                                value: true
                                            });
                                        }

                                        // Set the amount to apply (this updates the line amount)
                                        paymentRecord.setCurrentSublistValue({
                                            sublistId: 'apply',
                                            fieldId: 'amount',
                                            value: amountFloat
                                        });

                                        // Commit the line changes (DYNAMIC MODE REQUIRED)
                                        paymentRecord.commitLine({
                                            sublistId: 'apply'
                                        });

                                        log.debug('Applied payment to invoice', {
                                            line: line,
                                            invoiceId: invoiceId,
                                            amount: amountFloat
                                        });
                                        break;
                                    }
                                }

                                if (!foundInvoice) {
                                    log.error('Target invoice not found on apply sublist', {
                                        targetInvoiceId: invoiceId,
                                        applyLineCount: applyLineCount
                                    });
                                }

                            } catch (applyError) {
                                log.error('Error applying payment to invoice', {
                                    error: applyError.message,
                                    stack: applyError.stack,
                                    invoiceId: invoiceId,
                                    amount: amountFloat
                                });
                                // Continue without applying - payment will still be created
                                // but may have the full invoice amount applied instead of custom amount
                            }
                        }

                        // Save the Customer Payment
                        var paymentId = paymentRecord.save();

                        log.audit('Customer Payment Created Successfully', {
                            paymentId: paymentId,
                            customer: customerIdInt,
                            amount: amountFloat,
                            wfAuthNumber: wfAuthNumber,
                            invoiceNumber: invoiceNumber,
                            invoiceId: invoiceId,
                            isTransformed: isTransformed
                        });

                        // Get the transaction ID from the newly created payment
                        var paymentTranId = '';
                        try {
                            var savedPaymentRecord = record.load({
                                type: record.Type.CUSTOMER_PAYMENT,
                                id: paymentId,
                                isDynamic: false
                            });
                            paymentTranId = savedPaymentRecord.getValue('tranid') || paymentId;
                        } catch (loadError) {
                            log.error('Error loading payment record for tranid', loadError.message);
                            paymentTranId = paymentId; // Fallback to ID
                        }

                        // Redirect back to the same page with success message
                        redirect.toSuitelet({
                            scriptId: context.request.parameters.script,
                            deploymentId: context.request.parameters.deploy,
                            parameters: {
                                success: 'true',
                                paymentTranId: paymentTranId,
                                paymentAmount: amountFloat,
                                appliedInvoice: invoiceNumber || 'N/A'
                            }
                        });

                    } catch (fieldError) {
                        log.error('Error setting fields on Customer Payment', {
                            error: fieldError.message,
                            stack: fieldError.stack,
                            customer: customerIdInt,
                            amount: amountFloat
                        });
                        throw fieldError;
                    }
                } else if (action === 'email_sales_rep') {
                    // Email Sales Rep action
                    var salesRepId = context.request.parameters.salesrep;
                    var salesOrderId = context.request.parameters.salesorder;
                    var salesOrderNumber = context.request.parameters.salesordernumber;
                    var emailBody = context.request.parameters.emailBody;
                    var emailSubject = context.request.parameters.emailSubject;

                    log.debug('Emailing Sales Rep', {
                        salesRepId: salesRepId,
                        salesOrderId: salesOrderId,
                        salesOrderNumber: salesOrderNumber,
                        emailSubject: emailSubject
                    });

                    // Validate inputs
                    if (!salesRepId || !salesOrderId || !emailBody || !emailSubject) {
                        throw new Error('Missing required parameters');
                    }

                    // Parse and validate sales rep ID
                    var salesRepIdInt = parseInt(salesRepId, 10);
                    if (isNaN(salesRepIdInt) || salesRepIdInt <= 0) {
                        throw new Error('Invalid sales rep ID: ' + salesRepId);
                    }

                    try {
                        // Get current user ID as sender
                        var currentUser = runtime.getCurrentUser();
                        var authorId = currentUser.id;

                        // Build CC recipients: Employee 185 (manager) and current user
                        var ccRecipients = [185];
                        if (authorId !== salesRepIdInt && authorId !== 185) {
                            ccRecipients.push(authorId);
                        }
                        
                        // Send email to sales rep
                        email.send({
                            author: authorId,
                            recipients: salesRepIdInt,
                            cc: ccRecipients,
                            subject: emailSubject,
                            body: emailBody,
                            relatedRecords: {
                                transactionId: parseInt(salesOrderId, 10)
                            }
                        });

                        log.audit('Email sent to sales rep', {
                            salesRepId: salesRepIdInt,
                            ccRecipients: ccRecipients,
                            salesOrderId: salesOrderId,
                            subject: emailSubject
                        });

                        // Look up sales rep name for the note
                        var salesRepName = 'Sales Rep';
                        try {
                            var empLookup = search.lookupFields({
                                type: search.Type.EMPLOYEE,
                                id: salesRepIdInt,
                                columns: ['firstname', 'lastname']
                            });
                            if (empLookup) {
                                var firstName = empLookup.firstname || '';
                                var lastName = empLookup.lastname || '';
                                salesRepName = (firstName + ' ' + lastName).trim();
                                if (!salesRepName) {
                                    salesRepName = 'Sales Rep ID ' + salesRepIdInt;
                                }
                            }
                        } catch (lookupError) {
                            log.error('Error looking up sales rep name', lookupError.message);
                        }

                        // Create note to log the email
                        try {
                            var noteRecord = record.create({
                                type: 'note',
                                isDynamic: false
                            });

                            noteRecord.setValue({
                                fieldId: 'title',
                                value: 'Wells Fargo Email Sent'
                            });

                            var noteText = 'Short authorization email sent to ' + salesRepName + '.\n' +
                                          'Manager Mohamad Alkayal copied on email.';

                            noteRecord.setValue({
                                fieldId: 'note',
                                value: noteText
                            });

                            noteRecord.setValue({
                                fieldId: 'transaction',
                                value: parseInt(salesOrderId, 10)
                            });

                            var noteId = noteRecord.save();

                            log.audit('Note created after email', {
                                noteId: noteId,
                                salesOrderId: salesOrderId,
                                salesRepName: salesRepName
                            });
                        } catch (noteError) {
                            log.error('Error creating note after email', {
                                error: noteError.message,
                                salesOrderId: salesOrderId
                            });
                            // Don't fail the whole operation if note creation fails
                        }

                        // Redirect back with success message
                        redirect.toSuitelet({
                            scriptId: context.request.parameters.script,
                            deploymentId: context.request.parameters.deploy,
                            parameters: {
                                success: 'email_sent',
                                salesOrderNumber: salesOrderNumber
                            }
                        });

                    } catch (emailError) {
                        log.error('Error sending email to sales rep', {
                            error: emailError.message,
                            stack: emailError.stack,
                            salesRepId: salesRepIdInt,
                            salesOrderId: salesOrderId
                        });
                        throw emailError;
                    }

                } else if (action === 'create_note') {
                    // Log ALL parameters to debug
                    log.debug('Creating Note - All POST Parameters', JSON.stringify(context.request.parameters));

                    var transactionId = context.request.parameters.transactionId;
                    var noteText = context.request.parameters.noteText;

                    log.debug('Creating Note - Input Values', {
                        transactionId: transactionId,
                        noteText: noteText
                    });

                    // Validate inputs
                    if (!transactionId || !noteText) {
                        throw new Error('Missing required parameters: transactionId=' + transactionId + ', noteText=' + noteText);
                    }

                    // Parse and validate transaction ID
                    var transactionIdInt = parseInt(transactionId, 10);
                    if (isNaN(transactionIdInt) || transactionIdInt <= 0) {
                        throw new Error('Invalid transaction ID: ' + transactionId);
                    }

                    // Create note record
                    var noteRecord = record.create({
                        type: 'note',
                        isDynamic: false
                    });

                    try {
                        // Set note fields
                        noteRecord.setValue({
                            fieldId: 'title',
                            value: 'Wells Fargo Note'
                        });

                        noteRecord.setValue({
                            fieldId: 'note',
                            value: noteText
                        });

                        // Link to transaction using transaction field
                        noteRecord.setValue({
                            fieldId: 'transaction',
                            value: transactionIdInt
                        });

                        // Don't set notedate - let NetSuite default it to current date in account timezone
                        // This avoids timezone conversion issues where JavaScript Date objects
                        // can cause the date to shift when converted to Pacific Time

                        // Save the note
                        var noteId = noteRecord.save();

                        log.audit('Note Created Successfully', {
                            noteId: noteId,
                            transactionId: transactionIdInt,
                            noteText: noteText
                        });

                        // Redirect back to the same page with success message
                        redirect.toSuitelet({
                            scriptId: context.request.parameters.script,
                            deploymentId: context.request.parameters.deploy,
                            parameters: {
                                success: 'note_created'
                            }
                        });

                    } catch (noteError) {
                        log.error('Error creating note', {
                            error: noteError.message,
                            stack: noteError.stack,
                            transactionId: transactionIdInt,
                            noteText: noteText
                        });
                        throw noteError;
                    }
                }

            } catch (e) {
                log.error('Error in POST processing', {
                    error: e.message,
                    stack: e.stack,
                    action: context.request.parameters.action,
                    customer: context.request.parameters.customer,
                    amount: context.request.parameters.amount
                });

                // Redirect back with error message
                redirect.toSuitelet({
                    scriptId: context.request.parameters.script,
                    deploymentId: context.request.parameters.deploy,
                    parameters: {
                        error: 'Error processing request: ' + e.message
                    }
                });
            }
        }
    }

    /**
     * Gets the total of all applied customer deposits for a Sales Order
     * @param {number} salesOrderId - The Sales Order internal ID
     * @returns {number} The total amount of customer deposits applied
     */
    function getAppliedCustomerDeposits(salesOrderId) {
        if (!salesOrderId) {
            return 0;
        }

        try {
            var depositSearch = search.create({
                type: search.Type.CUSTOMER_DEPOSIT,
                filters: [
                    ['appliedtotransaction.internalid', 'anyof', salesOrderId],
                    'AND',
                    ['mainline', 'is', 'T'],
                    'AND',
                    ['status', 'noneof', 'CustDep:V'] // Exclude voided deposits
                ],
                columns: ['total']
            });

            var total = 0;
            depositSearch.run().each(function(result) {
                var payment = parseFloat(result.getValue('total')) || 0;
                total += payment;
                return true; // Continue iterating
            });

            log.debug('Applied customer deposits retrieved', {
                salesOrderId: salesOrderId,
                total: total
            });

            return total;

        } catch (e) {
            log.error('Error getting applied customer deposits', {
                error: e.message,
                salesOrderId: salesOrderId
            });
            return 0;
        }
    }

    /**
     * Gets the transaction total from a Sales Order by internal ID
     * @param {number} salesOrderId - The Sales Order internal ID
     * @returns {number} The transaction total amount
     */
    function getSalesOrderTotal(salesOrderId) {
        if (!salesOrderId) {
            return 0;
        }

        try {
            var soRecord = record.load({
                type: record.Type.SALES_ORDER,
                id: salesOrderId,
                isDynamic: false
            });

            var total = soRecord.getValue({
                fieldId: 'total'
            });

            log.debug('Sales Order total retrieved', {
                salesOrderId: salesOrderId,
                total: total
            });

            return parseFloat(total) || 0;

        } catch (e) {
            log.error('Error getting Sales Order total', {
                error: e.message,
                salesOrderId: salesOrderId
            });
            return 0;
        }
    }

    /**
     * Looks up the fulfilling location from a department record
     * @param {number} departmentId - The department ID to lookup
     * @returns {number} The fulfilling location ID
     */
    function lookupFulfillingLocation(departmentId) {
        try {
            log.debug('Looking up fulfilling location', 'Department ID: ' + departmentId);

            var departmentRecord = record.load({
                type: record.Type.DEPARTMENT,
                id: departmentId,
                isDynamic: false
            });

            var fulfillingLocationId = departmentRecord.getValue({
                fieldId: 'custrecord_bas_fulfilling_location'
            });

            if (fulfillingLocationId) {
                log.debug('Fulfilling location found', {
                    departmentId: departmentId,
                    fulfillingLocationId: fulfillingLocationId
                });
                return parseInt(fulfillingLocationId, 10);
            } else {
                log.debug('No fulfilling location found, using default', 'Department ID: ' + departmentId);
                return 1; // Default fallback
            }

        } catch (e) {
            log.error('Error looking up fulfilling location', {
                error: e.message,
                departmentId: departmentId
            });
            return 1; // Default fallback
        }
    }

    /**
     * Gets existing deposit links from Wells Fargo Authorization record
     * @param {string} wfAuthId - Wells Fargo Authorization ID
     * @returns {Array} Array of existing deposit IDs
     */
    function getExistingDepositLinks(wfAuthId) {
        try {
            var wfAuthRecord = record.load({
                type: 'customrecord_bas_wf_auth',
                id: wfAuthId,
                isDynamic: false
            });

            var existingLinks = wfAuthRecord.getValue('custrecord_customer_deposit_link');

            if (!existingLinks) {
                return [];
            }

            // Handle both single value and array
            if (Array.isArray(existingLinks)) {
                return existingLinks;
            } else {
                return [existingLinks];
            }

        } catch (e) {
            log.error('Error getting existing deposit links', {
                error: e.message,
                wfAuthId: wfAuthId
            });
            return [];
        }
    }

    /**
     * Appends new deposit ID to existing multiple select values
     * @param {Array} existingLinks - Array of existing deposit IDs
     * @param {string} newDepositId - New deposit ID to append
     * @returns {Array} Updated array of deposit IDs
     */
    function appendDepositToMultipleSelect(existingLinks, newDepositId) {
        var updatedLinks = existingLinks.slice(); // Create copy

        // Append new deposit ID to the array
        updatedLinks.push(parseInt(newDepositId, 10));

        return updatedLinks;
    }

    /**
     * Batch loads the most recent note for multiple transactions at once
     * @param {Array<number>} transactionIds - Array of transaction internal IDs
     * @returns {Object} Map of transactionId -> note object
     */
    function batchLoadMostRecentNotes(transactionIds) {
        var noteMap = {};
        
        if (!transactionIds || transactionIds.length === 0) {
            return noteMap;
        }

        try {
            // Build IN clause with transaction IDs
            var idList = transactionIds.join(',');
            
            // Use SuiteQL with window function to get only the most recent note per transaction
            var sql = 'SELECT * FROM (' +
                'SELECT ' +
                'TransactionNote.Transaction, ' +
                'TransactionNote.ID, ' +
                'TransactionNote.NoteDate, ' +
                'TransactionNote.Note, ' +
                'TransactionNote.Title, ' +
                '(Employee.FirstName || \' \' || Employee.LastName) AS Author, ' +
                'ROW_NUMBER() OVER (PARTITION BY TransactionNote.Transaction ORDER BY TransactionNote.NoteDate DESC) AS rn ' +
                'FROM TransactionNote ' +
                'INNER JOIN Employee ON (Employee.ID = TransactionNote.Author) ' +
                'WHERE TransactionNote.Transaction IN (' + idList + ') ' +
                ') WHERE rn = 1';

            var results = query.runSuiteQL({
                query: sql
            }).asMappedResults();

            log.debug('Batch loaded notes', {
                transactionCount: transactionIds.length,
                notesFound: results ? results.length : 0
            });

            // Build map of transaction ID to note
            if (results && results.length > 0) {
                for (var i = 0; i < results.length; i++) {
                    var result = results[i];
                    noteMap[result.transaction] = {
                        id: result.id || '',
                        notedate: result.notedate || '',
                        note: result.note || '',
                        title: result.title || '',
                        author: result.author || ''
                    };
                }
            }

            return noteMap;
        } catch (e) {
            log.error('Error batch loading notes', {
                error: e.message,
                stack: e.stack
            });
            return noteMap;
        }
    }

    /**
     * Gets the most recent note for a given transaction using SuiteQL TransactionNote table
     * @param {number} transactionId - The transaction internal ID
     * @returns {Object|null} Note object with notedate, note text, author, and note ID, or null if no notes found
     * @deprecated Use batchLoadMostRecentNotes for better performance
     */
    function getMostRecentNote(transactionId) {
        if (!transactionId) {
            return null;
        }

        try {
            // Use SuiteQL to query TransactionNote table with Employee join for author name
            var sql = 'SELECT ' +
                'TransactionNote.ID, ' +
                'TransactionNote.NoteDate, ' +
                'TransactionNote.Note, ' +
                'TransactionNote.Title, ' +
                '(Employee.FirstName || \' \' || Employee.LastName) AS Author ' +
                'FROM TransactionNote ' +
                'INNER JOIN Employee ON (Employee.ID = TransactionNote.Author) ' +
                'WHERE TransactionNote.Transaction = ? ' +
                'ORDER BY TransactionNote.NoteDate DESC';

            var results = query.runSuiteQL({
                query: sql,
                params: [transactionId]
            }).asMappedResults();

            log.debug('getMostRecentNote query results', {
                transactionId: transactionId,
                resultsCount: results ? results.length : 0,
                firstResult: results && results.length > 0 ? JSON.stringify(results[0]) : 'none'
            });

            if (results && results.length > 0) {
                return {
                    id: results[0].id || '',
                    notedate: results[0].notedate || '',
                    note: results[0].note || '',
                    title: results[0].title || '',
                    author: results[0].author || ''
                };
            }

            return null;
        } catch (e) {
            log.error('Error fetching note for transaction ' + transactionId, {
                error: e.message,
                stack: e.stack
            });
            return null;
        }
    }

    /**
     * Batch loads Wells Fargo Authorization records for multiple Sales Orders at once
     * @param {Array<number>} salesOrderIds - Array of Sales Order internal IDs
     * @returns {Object} Map of salesOrderId -> {records, total, authNumbers}
     */
    function batchLoadWellsFargoAuths(salesOrderIds) {
        var authMap = {};
        
        if (!salesOrderIds || salesOrderIds.length === 0) {
            return authMap;
        }

        try {
            // Search for all Wells Fargo Authorization records for all Sales Orders at once
            var wfAuthSearch = search.create({
                type: 'customrecord_bas_wf_auth',
                filters: [
                    ['custrecord_bas_wf_so_number', 'anyof', salesOrderIds]
                ],
                columns: [
                    'custrecord_bas_wf_so_number', // Sales Order
                    'name', // Record name
                    'custrecord26', // Authorization amount
                    'custrecord25'  // Authorization number
                ]
            });

            // Initialize empty result for each sales order
            for (var i = 0; i < salesOrderIds.length; i++) {
                authMap[salesOrderIds[i]] = {
                    records: [],
                    total: 0,
                    authNumbers: []
                };
            }

            wfAuthSearch.run().each(function(result) {
                var soId = result.getValue('custrecord_bas_wf_so_number');
                var recordName = result.getValue('name') || '';
                var amount = parseFloat(result.getValue('custrecord26')) || 0;
                var authNumber = result.getValue('custrecord25') || '';
                var recordId = result.id;

                if (soId && authMap[soId]) {
                    authMap[soId].records.push({
                        id: recordId,
                        name: recordName,
                        amount: amount,
                        authNumber: authNumber
                    });
                    authMap[soId].total += amount;
                    if (authNumber) {
                        authMap[soId].authNumbers.push(authNumber);
                    }
                }

                return true; // Continue iteration
            });

            log.debug('Batch loaded Wells Fargo Auths', {
                salesOrderCount: salesOrderIds.length,
                authsFound: Object.keys(authMap).filter(function(k) { return authMap[k].records.length > 0; }).length
            });

            return authMap;

        } catch (e) {
            log.error('Error batch loading Wells Fargo Auths', {
                error: e.message,
                stack: e.stack
            });
            return authMap;
        }
    }

    /**
     * Gets all Wells Fargo Authorization records linked to a Sales Order
     * @param {number} salesOrderId - The Sales Order internal ID
     * @returns {Object} Object containing WF auth records array and total amount
     * @deprecated Use batchLoadWellsFargoAuths for better performance
     */
    function getWellsFargoAuthsForSalesOrder(salesOrderId) {
        if (!salesOrderId) {
            return { records: [], total: 0, authNumbers: [] };
        }

        try {
            // Search for all Wells Fargo Authorization records linked to this Sales Order
            var wfAuthSearch = search.create({
                type: 'customrecord_bas_wf_auth',
                filters: [
                    ['custrecord_bas_wf_so_number', 'anyof', salesOrderId]
                ],
                columns: [
                    'name', // Record name
                    'custrecord26', // Authorization amount
                    'custrecord25'  // Authorization number
                ]
            });

            var wfAuthRecords = [];
            var totalAmount = 0;
            var authNumbers = [];

            wfAuthSearch.run().each(function(result) {
                var recordName = result.getValue('name') || '';
                var amount = parseFloat(result.getValue('custrecord26')) || 0;
                var authNumber = result.getValue('custrecord25') || '';
                var recordId = result.id;

                wfAuthRecords.push({
                    id: recordId,
                    name: recordName,
                    amount: amount,
                    authNumber: authNumber
                });

                totalAmount += amount;
                if (authNumber) {
                    authNumbers.push(authNumber);
                }

                return true; // Continue iteration
            });

            log.debug('Wells Fargo Auths retrieved for Sales Order', {
                salesOrderId: salesOrderId,
                recordCount: wfAuthRecords.length,
                totalAmount: totalAmount
            });

            return {
                records: wfAuthRecords,
                total: totalAmount,
                authNumbers: authNumbers
            };

        } catch (e) {
            log.error('Error getting Wells Fargo Auths for Sales Order', {
                error: e.message,
                stack: e.stack,
                salesOrderId: salesOrderId
            });
            return { records: [], total: 0, authNumbers: [] };
        }
    }

    /**
  * Builds HTML content containing both search results
  * @param {Object} context - NetSuite context object containing request/response
  * @returns {string} HTML content string
  */
    function buildSearchResultsHTML(context) {
        var html = '<div id="loadingOverlay" class="loading-overlay">' +
            '<div class="loading-content">' +
            '<div class="loading-spinner"></div>' +
            '<div id="loadingText" class="loading-text">Processing...</div>' +
            '</div>' +
            '</div>';
        
        html += '<style>' +
            // Reset NetSuite default styles and remove borders
            '.uir-page-title-secondline { border: none !important; margin: 0 !important; padding: 0 !important; }' +
            '.uir-record-type { border: none !important; }' +
            '.bglt { border: none !important; }' +
            '.smalltextnolink { border: none !important; }' +

            // Main container styling
            '.wells-fargo-container { margin: 0; padding: 8px 8px 20px 8px; border: none; background: transparent; position: relative; }' +

            // SOP Quick Link styling
            '.sop-link-container { position: absolute; top: 0; right: 0; z-index: 100; }' +
            '.sop-quick-link { display: inline-flex; align-items: center; padding: 10px 18px; background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); color: white; text-decoration: none; border-radius: 8px; font-weight: 600; font-size: 13px; box-shadow: 0 4px 6px rgba(76, 175, 80, 0.3), 0 1px 3px rgba(0, 0, 0, 0.1); transition: all 0.3s ease; border: 2px solid rgba(255, 255, 255, 0.2); }' +
            '.sop-quick-link:hover { background: linear-gradient(135deg, #45a049 0%, #4CAF50 100%); transform: translateY(-2px); box-shadow: 0 6px 12px rgba(76, 175, 80, 0.4), 0 2px 4px rgba(0, 0, 0, 0.15); text-decoration: none; color: white; border-color: rgba(255, 255, 255, 0.3); }' +
            '.sop-quick-link:active { transform: translateY(0px); box-shadow: 0 2px 4px rgba(76, 175, 80, 0.3), 0 1px 2px rgba(0, 0, 0, 0.1); }' +
            '.sop-quick-link svg { filter: drop-shadow(0 1px 2px rgba(0, 0, 0, 0.2)); }' +

            // Table styling
            'table.search-table { border-collapse: collapse; width: 100%; margin: 15px 0; border: 1px solid #ddd; background: white; }' +
            'table.search-table th, table.search-table td { border: 1px solid #ddd; padding: 5px; text-align: left; vertical-align: top; font-size: 11px; }' +
            'table.search-table th { background-color: #f8f9fa; font-weight: bold; color: #333; font-size: 11px; position: -webkit-sticky; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); }' +
            'table.search-table tr:nth-child(even) td { background-color: #f9f9f9; }' +
            'table.search-table tr:hover td { background-color: #e8f4f8; }' +

            // Search section wrapper (non-sticky)
            '.search-section-header { background: white; }' +
            '.search-title { font-size: 16px; font-weight: bold; margin: 25px 0 0 0; color: #333; padding: 10px 10px 6px 10px; border-bottom: 2px solid #4CAF50; }' +
            '.search-count { font-style: italic; color: #666; margin: 0; font-size: 12px; padding: 2px 10px 0 10px; }' +
            '.results-count { font-style: italic; color: #666; margin: 0; font-size: 12px; padding: 0 10px 8px 10px; }' +

            // Button styling
            '.action-btn { background-color: #4CAF50; color: white; padding: 6px 12px; border: none; cursor: pointer; border-radius: 4px; font-size: 11px; text-decoration: none; display: inline-block; transition: background-color 0.3s; }' +
            '.action-btn:hover { background-color: #45a049; text-decoration: none; }' +
            '.action-btn:disabled { background-color: #cccccc; cursor: not-allowed; }' +
            '.action-btn-secondary { background-color: #5b9bd5; color: white; padding: 5px 10px; border: none; cursor: pointer; border-radius: 3px; font-size: 10px; text-decoration: none; display: inline-block; transition: background-color 0.3s; }' +
            '.action-btn-secondary:hover { background-color: #4a8bc2; text-decoration: none; }' +
            '.action-cell { text-align: center; white-space: nowrap; padding: 4px; }' +

            // Note-related styling
            '.has-note td { background-color: #fff9e6 !important; }' +
            '.has-note td:first-child { border-left: 3px solid #ffd966 !important; }' +
            '.has-note:hover td { background-color: #fff4cc !important; }' +
            '.note-cell { max-width: 300px; font-size: 11px; word-wrap: break-word; }' +

            // Deposit validation styling
            '.deposit-validation { font-size: inherit; line-height: inherit; white-space: pre-line; }' +
            '.validation-match { color: #155724; background-color: #d4edda; padding: 2px 4px; border-radius: 3px; font-weight: bold; }' +
            '.validation-short { color: #721c24; background-color: #f8d7da; padding: 2px 4px; border-radius: 3px; font-weight: bold; }' +
            '.validation-over { color: #856404; background-color: #fff3cd; padding: 2px 4px; border-radius: 3px; font-weight: bold; }' +

            // Note dialog styling
            '.note-dialog-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.5); z-index: 10000; justify-content: center; align-items: center; }' +
            '.note-dialog-overlay.active { display: flex; }' +
            '.note-dialog { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3); min-width: 400px; max-width: 600px; }' +
            '.note-dialog h3 { margin-top: 0; margin-bottom: 15px; color: #333; }' +
            '.note-dialog textarea { width: 100%; min-height: 100px; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-family: inherit; font-size: 13px; resize: vertical; }' +
            '.note-dialog-buttons { margin-top: 15px; text-align: right; }' +
            '.note-dialog-buttons button { margin-left: 10px; }' +

            // Message styling
            '.success-msg { background-color: #d4edda; color: #155724; padding: 12px; border: 1px solid #c3e6cb; border-radius: 6px; margin: 15px 0; font-size: 13px; }' +
            '.error-msg { background-color: #f8d7da; color: #721c24; padding: 12px; border: 1px solid #f5c6cb; border-radius: 6px; margin: 15px 0; font-size: 13px; }' +

            // Loading overlay styling
            '.loading-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0, 0, 0, 0.5); z-index: 9999; justify-content: center; align-items: center; }' +
            '.loading-overlay.active { display: flex; }' +
            '.loading-content { background-color: white; padding: 30px; border-radius: 8px; text-align: center; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); }' +
            '.loading-spinner { border: 4px solid #f3f3f3; border-top: 4px solid #4CAF50; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto 15px; }' +
            '@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
            '.loading-text { font-size: 14px; color: #333; font-weight: bold; }' +

            // Hidden data containers
            '.hidden-data { display: none; }' +

            // Ensure clean container
            'body, html { margin: 0; padding: 0; }' +
            '</style>';

        // Add inline JavaScript functions
        html += '<script>' +
            'function refreshPage() { window.location.reload(); }' +

            // Show loading overlay
            'function showLoading(message) {' +
            '    var overlay = document.getElementById("loadingOverlay");' +
            '    var text = document.getElementById("loadingText");' +
            '    if (overlay && text) {' +
            '        text.textContent = message || "Processing...";' +
            '        overlay.className = "loading-overlay active";' +
            '    }' +
            '}' +

            // Hide loading overlay
            'function hideLoading() {' +
            '    var overlay = document.getElementById("loadingOverlay");' +
            '    if (overlay) {' +
            '        overlay.className = "loading-overlay";' +
            '    }' +
            '}' +

            // Prompt for deposit amount and submit
            'function promptAndSubmitDeposit(dataId, defaultAmount) {' +
            '    try {' +
            '        var amount = window.prompt("Enter deposit amount:", defaultAmount);' +
            '        if (amount === null) {' +
            '            return;' +
            '        }' +
            '        var numAmount = parseFloat(amount);' +
            '        if (isNaN(numAmount) || numAmount <= 0) {' +
            '            alert("Please enter a valid amount greater than zero");' +
            '            return;' +
            '        }' +
            '        numAmount = Math.round(numAmount * 100) / 100;' +
            '        ' +
            '        var dataContainer = document.getElementById(dataId);' +
            '        if (!dataContainer) {' +
            '            alert("Error: Data container not found - ID: " + dataId);' +
            '            return;' +
            '        }' +
            '        ' +
            '        showLoading("Creating customer deposit...");' +
            '        ' +
            '        var form = document.createElement("form");' +
            '        form.method = "POST";' +
            '        form.action = window.location.href;' +
            '        ' +
            '        var inputs = dataContainer.getElementsByTagName("input");' +
            '        for (var i = 0; i < inputs.length; i++) {' +
            '            var input = inputs[i].cloneNode(true);' +
            '            if (input.name === "amount") {' +
            '                input.value = numAmount.toFixed(2);' +
            '            }' +
            '            form.appendChild(input);' +
            '        }' +
            '        ' +
            '        document.body.appendChild(form);' +
            '        form.submit();' +
            '    } catch (e) {' +
            '        hideLoading();' +
            '        alert("Error: " + e.message);' +
            '    }' +
            '}' +

            // Prompt for payment amount and submit
            'function promptAndSubmitPayment(dataId, defaultAmount) {' +
            '    try {' +
            '        var amount = window.prompt("Enter payment amount:", defaultAmount);' +
            '        if (amount === null) {' +
            '            return;' +
            '        }' +
            '        var numAmount = parseFloat(amount);' +
            '        if (isNaN(numAmount) || numAmount <= 0) {' +
            '            alert("Please enter a valid amount greater than zero");' +
            '            return;' +
            '        }' +
            '        numAmount = Math.round(numAmount * 100) / 100;' +
            '        ' +
            '        var dataContainer = document.getElementById(dataId);' +
            '        if (!dataContainer) {' +
            '            alert("Error: Data container not found - ID: " + dataId);' +
            '            return;' +
            '        }' +
            '        ' +
            '        showLoading("Creating customer payment...");' +
            '        ' +
            '        var form = document.createElement("form");' +
            '        form.method = "POST";' +
            '        form.action = window.location.href;' +
            '        ' +
            '        var inputs = dataContainer.getElementsByTagName("input");' +
            '        for (var i = 0; i < inputs.length; i++) {' +
            '            var input = inputs[i].cloneNode(true);' +
            '            if (input.name === "amount") {' +
            '                input.value = numAmount.toFixed(2);' +
            '            }' +
            '            form.appendChild(input);' +
            '        }' +
            '        ' +
            '        document.body.appendChild(form);' +
            '        form.submit();' +
            '    } catch (e) {' +
            '        hideLoading();' +
            '        alert("Error: " + e.message);' +
            '    }' +
            '}' +

            // Show note dialog
            'function promptAndCreateNote(dataId) {' +
            '    var dialog = document.getElementById("noteDialog");' +
            '    var textarea = document.getElementById("noteTextarea");' +
            '    var currentDataId = document.getElementById("currentNoteDataId");' +
            '    ' +
            '    if (dialog && textarea && currentDataId) {' +
            '        currentDataId.value = dataId;' +
            '        textarea.value = "";' +
            '        dialog.className = "note-dialog-overlay active";' +
            '        textarea.focus();' +
            '    }' +
            '}' +

            // Hide note dialog
            'function closeNoteDialog() {' +
            '    var dialog = document.getElementById("noteDialog");' +
            '    if (dialog) {' +
            '        dialog.className = "note-dialog-overlay";' +
            '    }' +
            '}' +

            // Submit note
            'function submitNote() {' +
            '    try {' +
            '        var textarea = document.getElementById("noteTextarea");' +
            '        var dataId = document.getElementById("currentNoteDataId").value;' +
            '        var noteText = textarea.value.trim();' +
            '        ' +
            '        if (!noteText) {' +
            '            alert("Please enter a note");' +
            '            return;' +
            '        }' +
            '        ' +
            '        closeNoteDialog();' +
            '        showLoading("Creating note...");' +
            '        ' +
            '        var dataContainer = document.getElementById(dataId);' +
            '        if (!dataContainer) {' +
            '            hideLoading();' +
            '            alert("Error: Data container not found");' +
            '            return;' +
            '        }' +
            '        ' +
            '        console.log("Data container ID:", dataId);' +
            '        console.log("Data container HTML:", dataContainer.innerHTML);' +
            '        ' +
            '        var form = document.createElement("form");' +
            '        form.method = "POST";' +
            '        form.action = window.location.href;' +
            '        ' +
            '        var inputs = dataContainer.getElementsByTagName("input");' +
            '        console.log("Found", inputs.length, "inputs in container");' +
            '        for (var i = 0; i < inputs.length; i++) {' +
            '            console.log("Input", i, "- name:", inputs[i].name, "value:", inputs[i].value);' +
            '            form.appendChild(inputs[i].cloneNode(true));' +
            '        }' +
            '        ' +
            '        var noteInput = document.createElement("input");' +
            '        noteInput.type = "hidden";' +
            '        noteInput.name = "noteText";' +
            '        noteInput.value = noteText;' +
            '        form.appendChild(noteInput);' +
            '        ' +
            '        document.body.appendChild(form);' +
            '        form.submit();' +
            '    } catch (e) {' +
            '        hideLoading();' +
            '        alert("Error: " + e.message);' +
            '    }' +
            '}' +

            // Email Sales Rep for SHORT authorizations
            'function emailSalesRep(dataId) {' +
            '    try {' +
            '        var container = document.getElementById(dataId);' +
            '        if (!container) {' +
            '            alert("Error: Data container not found");' +
            '            return;' +
            '        }' +
            '        ' +
            '        var salesOrderNumber = container.querySelector("[name=\\"salesordernumber\\"]").value;' +
            '        var customerName = container.querySelector("[name=\\"customername\\"]").value;' +
            '        var sellingLocation = container.querySelector("[name=\\"sellinglocation\\"]").value;' +
            '        var soTotal = container.querySelector("[name=\\"sototal\\"]").value;' +
            '        var required = container.querySelector("[name=\\"required\\"]").value;' +
            '        var priorDeposits = container.querySelector("[name=\\"priordeposits\\"]").value;' +
            '        var wfAmount = container.querySelector("[name=\\"amount\\"]").value;' +
            '        var afterProcessing = container.querySelector("[name=\\"afterprocessing\\"]").value;' +
            '        var variance = container.querySelector("[name=\\"variance\\"]").value;' +
            '        ' +
            '        var defaultBody = "Wells Fargo Kitchen Works Short Authorization\\n\\n" +' +
            '                         salesOrderNumber + "\\n" +' +
            '                         "Customer: " + customerName + "\\n" +' +
            '                         "Selling Location: " + sellingLocation + "\\n\\n" +' +
            '                         "DEPOSIT VALIDATION:\\n" +' +
            '                         "SO Total: " + soTotal + "\\n" +' +
            '                         "Required (50% CD): " + required + "\\n" +' +
            '                         "Prior CDs: " + priorDeposits + "\\n" +' +
            '                         "WF To Be Processed: " + wfAmount + "\\n" +' +
            '                         "Total CDs After Processing: " + afterProcessing + "\\n\\n" +' +
            '                         "⚠️ SHORT " + Math.abs(parseFloat(variance)).toFixed(2) + "\\n\\n" +' +
            '                         "Please review this authorization and follow up on the deposit shortage.\\n\\n" +' +
            '                         "REMINDER: Authorizations should be for the FULL balance of the Sales Order, allowing us to charge 50% as a Customer Deposit and the remaining balance at the time of invoicing.";' +
            '        ' +
            '        var dialog = document.getElementById("emailDialog");' +
            '        var textarea = document.getElementById("emailTextarea");' +
            '        var subjectField = document.getElementById("emailSubject");' +
            '        var currentDataId = document.getElementById("currentEmailDataId");' +
            '        ' +
            '        if (dialog && textarea && subjectField && currentDataId) {' +
            '            currentDataId.value = dataId;' +
            '            textarea.value = defaultBody;' +
            '            subjectField.value = "WF Short Auth - Kitchen Works - " + customerName;' +
            '            dialog.className = "note-dialog-overlay active";' +
            '            textarea.focus();' +
            '        }' +
            '    } catch (e) {' +
            '        alert("Error: " + e.message);' +
            '    }' +
            '}' +

            // Close email dialog
            'function closeEmailDialog() {' +
            '    var dialog = document.getElementById("emailDialog");' +
            '    if (dialog) {' +
            '        dialog.className = "note-dialog-overlay";' +
            '    }' +
            '}' +

            // Submit email
            'function submitEmail() {' +
            '    try {' +
            '        var textarea = document.getElementById("emailTextarea");' +
            '        var subjectField = document.getElementById("emailSubject");' +
            '        var dataId = document.getElementById("currentEmailDataId").value;' +
            '        var emailBody = textarea.value.trim();' +
            '        var emailSubject = subjectField.value.trim();' +
            '        ' +
            '        if (!emailBody) {' +
            '            alert("Please enter an email body");' +
            '            return;' +
            '        }' +
            '        ' +
            '        if (!emailSubject) {' +
            '            alert("Please enter an email subject");' +
            '            return;' +
            '        }' +
            '        ' +
            '        closeEmailDialog();' +
            '        showLoading("Sending email...");' +
            '        ' +
            '        var dataContainer = document.getElementById(dataId);' +
            '        if (!dataContainer) {' +
            '            hideLoading();' +
            '            alert("Error: Data container not found");' +
            '            return;' +
            '        }' +
            '        ' +
            '        var form = document.createElement("form");' +
            '        form.method = "POST";' +
            '        form.action = window.location.href;' +
            '        ' +
            '        var actionInput = document.createElement("input");' +
            '        actionInput.type = "hidden";' +
            '        actionInput.name = "action";' +
            '        actionInput.value = "email_sales_rep";' +
            '        form.appendChild(actionInput);' +
            '        ' +
            '        var inputs = dataContainer.getElementsByTagName("input");' +
            '        for (var i = 0; i < inputs.length; i++) {' +
            '            if (inputs[i].name !== "action") {' +
            '                form.appendChild(inputs[i].cloneNode(true));' +
            '            }' +
            '        }' +
            '        ' +
            '        var emailBodyInput = document.createElement("input");' +
            '        emailBodyInput.type = "hidden";' +
            '        emailBodyInput.name = "emailBody";' +
            '        emailBodyInput.value = emailBody;' +
            '        form.appendChild(emailBodyInput);' +
            '        ' +
            '        var emailSubjectInput = document.createElement("input");' +
            '        emailSubjectInput.type = "hidden";' +
            '        emailSubjectInput.name = "emailSubject";' +
            '        emailSubjectInput.value = emailSubject;' +
            '        form.appendChild(emailSubjectInput);' +
            '        ' +
            '        document.body.appendChild(form);' +
            '        form.submit();' +
            '    } catch (e) {' +
            '        hideLoading();' +
            '        alert("Error: " + e.message);' +
            '    }' +
            '}' +
            '</script>';

        // Add note dialog HTML
        html += '<div id="noteDialog" class="note-dialog-overlay">' +
            '<div class="note-dialog">' +
            '<h3>Create Note</h3>' +
            '<textarea id="noteTextarea" placeholder="Enter your note here..."></textarea>' +
            '<input type="hidden" id="currentNoteDataId" value="">' +
            '<div class="note-dialog-buttons">' +
            '<button type="button" class="action-btn" style="background: #666;" onclick="closeNoteDialog()">Cancel</button>' +
            '<button type="button" class="action-btn" onclick="submitNote()">Save Note</button>' +
            '</div>' +
            '</div>' +
            '</div>';

        // Add email dialog HTML
        html += '<div id="emailDialog" class="note-dialog-overlay">' +
            '<div class="note-dialog" style="max-width: 600px;">' +
            '<h3>Send Email to Sales Rep</h3>' +
            '<label style="display: block; margin-bottom: 5px; font-weight: bold;">Subject:</label>' +
            '<input type="text" id="emailSubject" style="width: 100%; padding: 8px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px;" />' +
            '<label style="display: block; margin-bottom: 5px; font-weight: bold;">Message:</label>' +
            '<textarea id="emailTextarea" placeholder="Enter email message..." style="height: 300px;"></textarea>' +
            '<input type="hidden" id="currentEmailDataId" value="">' +
            '<div class="note-dialog-buttons">' +
            '<button type="button" class="action-btn" style="background: #666;" onclick="closeEmailDialog()">Cancel</button>' +
            '<button type="button" class="action-btn" onclick="submitEmail()">Send Email</button>' +
            '</div>' +
            '</div>' +
            '</div>';

        // Main container
        html += '<div class="wells-fargo-container">';

        // SOP Quick Link (top right)
        html += '<div class="sop-link-container">';
        html += '<a href="https://8289753.app.netsuite.com/app/site/hosting/scriptlet.nl?script=3923&deploy=1#15" target="_blank" class="sop-quick-link">';
        html += '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 6px; vertical-align: middle;">';
        html += '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>';
        html += '<polyline points="14 2 14 8 20 8"></polyline>';
        html += '<line x1="16" y1="13" x2="8" y2="13"></line>';
        html += '<line x1="16" y1="17" x2="8" y2="17"></line>';
        html += '<polyline points="10 9 9 9 8 9"></polyline>';
        html += '</svg>';
        html += '<span>Kitchens SOP Quick Link</span>';
        html += '</a>';
        html += '</div>';

        // Show success/error messages with XSS protection
        if (context && context.request.parameters.success) {
            html += '<div class="success-msg">';

            if (context.request.parameters.success === 'email_sent') {
                // Email sent success message
                html += '<strong>Email Sent Successfully</strong><br>';
                html += 'Sales rep has been notified about the short authorization for ' + escapeHtml(context.request.parameters.salesOrderNumber || 'the sales order') + '.';
            } else if (context.request.parameters.success === 'note_created') {
                // Note success message
                html += '<strong>Note Created Successfully</strong><br>';
                html += 'A new note has been added to the transaction.';
            } else if (context.request.parameters.depositTranId) {
                // Deposit success message
                html += '<strong>Customer Deposit Created Successfully and Wells Fargo Authorization Record Updated</strong><br>';
                html += 'Customer Deposit: ' + escapeHtml(context.request.parameters.depositTranId || 'Unknown') + '<br>';
                html += 'Wells Fargo Authorization: ' + escapeHtml(context.request.parameters.wfAuthName || 'Unknown');
            } else if (context.request.parameters.paymentTranId) {
                // Payment success message
                html += '<strong>Customer Payment Created Successfully</strong><br>';
                html += 'Customer Payment: ' + escapeHtml(context.request.parameters.paymentTranId || 'Unknown') + '<br>';
                html += 'Amount: $' + escapeHtml(context.request.parameters.paymentAmount || 'Unknown') + '<br>';
                html += 'Applied to Invoice: ' + escapeHtml(context.request.parameters.appliedInvoice || 'N/A');
            }

            html += '</div>';
        }
        if (context && context.request.parameters.error) {
            html += '<div class="error-msg"><strong>Error:</strong> ' + escapeHtml(context.request.parameters.error) + '</div>';
        }

        // First Search: Wells Fargo Sales Order Customer Deposits
        html += buildSearchSection(
            'Kitchen Works Materials: Sales Order 50% Deposit Processing',
            'Saved Search Data: BAS Wells Fargo Sales Order Customer Deposits To Be Charged',
            'customsearch_bas_wells_fargo_so_cd',
            11,
            'deposit',
            true  // Enable notes for deposit table
        );

        // Second Search: A/R Aging (Wells Fargo Financing)
        html += buildSearchSection(
            'Open A/R Invoice / Credit Memo Processing',
            'Saved Search Data: BAS A/R Aging (Wells Fargo Financing)',
            'customsearch5263',
            16,
            'payment',
            true  // Enable notes for payment table
        );

        // Close main container
        html += '</div>';

        // Hide initial loading overlay after page and content are fully loaded
        html += '<script>' +
            'document.addEventListener("DOMContentLoaded", function() {' +
            '    setTimeout(function() {' +
            '        var overlay = document.getElementById("loadingOverlay");' +
            '        if (overlay) {' +
            '            overlay.className = "loading-overlay";' +
            '        }' +
            '    }, 100);' +
            '});' +
            '</script>';

        return html;
    }

    /**
     * Builds a complete search section with sticky header
     * @param {string} title - The section title
     * @param {string} searchInfo - The search information text
     * @param {string} searchId - The saved search ID
     * @param {number} expectedColumns - Expected number of columns to display
     * @param {string} actionType - Type of action ('deposit' or 'payment')
     * @param {boolean} enableNotes - Whether to show notes column and functionality
     * @returns {string} HTML string for the complete section
     */
    function buildSearchSection(title, searchInfo, searchId, expectedColumns, actionType, enableNotes) {
        var html = '';

        // Start sticky header wrapper
        html += '<div class="search-section-header">';
        html += '<div class="search-title">' + escapeHtml(title) + '</div>';
        html += '<div class="search-count">' + escapeHtml(searchInfo) + '</div>';

        // Add results count inside the sticky header
        try {
            var savedSearch = search.load({ id: searchId });
            var searchResults = savedSearch.run();
            var resultsRange = searchResults.getRange({ start: 0, end: 1000 });
            html += '<div class="results-count">Results: ' + resultsRange.length + '</div>';
        } catch (e) {
            html += '<div class="results-count">Results: Error loading</div>';
        }

        html += '</div>'; // Close sticky header wrapper

        // Add the table
        html += buildSearchTable(searchId, expectedColumns, actionType, enableNotes);

        return html;
    }

    /**
     * Builds HTML table for a specific saved search
     * @param {string} searchId - The saved search ID
     * @param {number} expectedColumns - Expected number of columns to display
     * @param {string} actionType - Type of action ('deposit' or 'payment')
     * @param {boolean} enableNotes - Whether to show notes column and functionality
     * @returns {string} HTML table string
     */
    function buildSearchTable(searchId, expectedColumns, actionType, enableNotes) {
        try {
            var savedSearch = search.load({
                id: searchId
            });

            var searchResults = savedSearch.run();
            var resultsRange = searchResults.getRange({
                start: 0,
                end: 1000
            });

            if (resultsRange.length === 0) {
                return '<div style="padding: 10px; font-style: italic; color: #666;">No results found</div>';
            }

            // BATCH LOAD OPTIMIZATION: Collect all transaction IDs and Sales Order IDs upfront
            var transactionIds = [];
            var salesOrderIds = [];
            
            if (enableNotes || actionType === 'payment') {
                for (var i = 0; i < resultsRange.length; i++) {
                    var tempData = extractRowData(resultsRange[i], actionType);
                    
                    // Collect transaction IDs for notes
                    if (enableNotes) {
                        var txnId = (actionType === 'deposit') ? tempData.salesOrderId : tempData.invoiceId;
                        if (txnId) {
                            transactionIds.push(txnId);
                        }
                    }
                    
                    // Collect Sales Order IDs for WF auth lookup (payment rows only)
                    if (actionType === 'payment' && tempData.salesOrderId) {
                        salesOrderIds.push(tempData.salesOrderId);
                    }
                }
            }
            
            // Batch load all notes and WF auths at once
            var noteMap = enableNotes && transactionIds.length > 0 ? batchLoadMostRecentNotes(transactionIds) : {};
            var wfAuthMap = salesOrderIds.length > 0 ? batchLoadWellsFargoAuths(salesOrderIds) : {};
            
            log.debug('Batch load complete', {
                actionType: actionType,
                rowCount: resultsRange.length,
                notesLoaded: Object.keys(noteMap).length,
                wfAuthsLoaded: Object.keys(wfAuthMap).length
            });

            var html = '<table class="search-table">';

            // Build header row
            html += '<thead><tr>';
            html += '<th>Action</th>';
            
            // For deposit action type, we need to insert Deposit Validation header
            var colsToDisplay = (actionType === 'deposit') ? expectedColumns - 1 : expectedColumns;
            
            for (var col = 0; col < colsToDisplay; col++) {
                // For deposit table, insert Deposit Validation header after column 7 (Selling Location)
                if (actionType === 'deposit' && col === 8) {
                    html += '<th>Deposit Validation</th>';
                }
                
                try {
                    var columnLabel = resultsRange[0].columns[col] ?
                        (resultsRange[0].columns[col].label || 'Column ' + (col + 1)) :
                        'Column ' + (col + 1);
                    
                    // Shorten column names to reduce width
                    var displayLabel = columnLabel;
                    if (columnLabel === 'Wells Fargo Authorization Record ID') {
                        displayLabel = 'WF Record ID';
                    } else if (columnLabel === 'Wells Fargo Authorization #') {
                        displayLabel = 'WF Auth #';
                    } else if (columnLabel === 'Wells Fargo Authorization Amount ($)') {
                        displayLabel = 'WF Auth Amount ($)';
                    }
                    
                    html += '<th>' + escapeHtml(displayLabel) + '</th>';
                } catch (e) {
                    html += '<th>Column ' + (col + 1) + '</th>';
                }
            }
            // Add User Notes column header if notes are enabled
            if (enableNotes) {
                html += '<th>User Notes</th>';
            }
            html += '</tr></thead>';

            // Build data rows
            html += '<tbody>';
            for (var i = 0; i < resultsRange.length; i++) {
                // Check if this row has notes (for highlighting) - works for both deposit and payment
                var rowClass = '';
                var recentNote = null;
                if (enableNotes) {
                    var tempData = extractRowData(resultsRange[i], actionType);
                    var transactionId = (actionType === 'deposit') ? tempData.salesOrderId : tempData.invoiceId;
                    
                    // Look up pre-loaded note from batch map
                    if (transactionId && noteMap[transactionId]) {
                        recentNote = noteMap[transactionId];
                        rowClass = 'has-note';
                    }
                }

                html += '<tr' + (rowClass ? ' class="' + rowClass + '"' : '') + '>';

                if (actionType === 'deposit') {
                    var mappedData = mapDepositColumns(resultsRange[i]);
                    var dataId = 'deposit_data_' + i;
                    var noteDataId = 'note_data_deposit_' + i;
                    
                    // Calculate validation data
                    var soTotal = parseFloat(mappedData.salesOrderTotal) || 0;
                    var required = soTotal / 2;
                    var priorDeposits = parseFloat(mappedData.customerDepositTotal) || 0;
                    var wfCharge = parseFloat(mappedData.amount) || 0;
                    var afterProcessing = priorDeposits + wfCharge;
                    var variance = afterProcessing - required;
                    var isShort = variance < 0;
                    
                    // Get display values for email
                    var salesOrderNumber = '';
                    var customerName = '';
                    var sellingLocation = '';
                    
                    for (var col = 0; col < resultsRange[i].columns.length; col++) {
                        var columnLabel = resultsRange[i].columns[col] ? (resultsRange[i].columns[col].label || '') : '';
                        var cellValue = resultsRange[i].getValue(resultsRange[i].columns[col]) || '';
                        var textValue = resultsRange[i].getText(resultsRange[i].columns[col]);
                        if (textValue && textValue !== 'undefined') {
                            cellValue = textValue;
                        }
                        
                        if (columnLabel === 'Sales Order #') {
                            salesOrderNumber = cellValue;
                        } else if (columnLabel === 'Customer') {
                            customerName = cellValue;
                        } else if (columnLabel === 'Selling Location') {
                            sellingLocation = cellValue;
                        }
                    }

                    html += '<td class="action-cell">';

                    // Hidden data container for deposit (NO FORM TAG - just a div with data)
                    html += '<div id="' + dataId + '" class="hidden-data">';
                    html += '<input type="hidden" name="action" value="create_deposit">';
                    html += '<input type="hidden" name="customer" value="' + escapeHtml(mappedData.customerId) + '">';
                    html += '<input type="hidden" name="salesorder" value="' + escapeHtml(mappedData.salesOrderId) + '">';
                    html += '<input type="hidden" name="amount" value="' + escapeHtml(mappedData.amount) + '">';
                    html += '<input type="hidden" name="wfAuthId" value="' + escapeHtml(mappedData.wfAuthId) + '">';
                    html += '<input type="hidden" name="location" value="' + escapeHtml(mappedData.location) + '">';
                    html += '<input type="hidden" name="wfAuthNumber" value="' + escapeHtml(mappedData.wfAuthNumber) + '">';
                    html += '<input type="hidden" name="salesrep" value="' + escapeHtml(mappedData.salesRep) + '">';
                    html += '<input type="hidden" name="salesordernumber" value="' + escapeHtml(salesOrderNumber) + '">';
                    html += '<input type="hidden" name="customername" value="' + escapeHtml(customerName) + '">';
                    html += '<input type="hidden" name="sellinglocation" value="' + escapeHtml(sellingLocation) + '">';
                    html += '<input type="hidden" name="sototal" value="' + soTotal.toFixed(2) + '">';
                    html += '<input type="hidden" name="required" value="' + required.toFixed(2) + '">';
                    html += '<input type="hidden" name="priordeposits" value="' + priorDeposits.toFixed(2) + '">';
                    html += '<input type="hidden" name="afterprocessing" value="' + afterProcessing.toFixed(2) + '">';
                    html += '<input type="hidden" name="variance" value="' + variance.toFixed(2) + '">';
                    html += '</div>';

                    // Primary action: Create Deposit button
                    html += '<button type="button" class="action-btn" onclick="promptAndSubmitDeposit(\'' + dataId + '\', \'' + escapeHtml(mappedData.amount) + '\')">Create Deposit</button>';

                    // Email Sales Rep button - only for SHORT rows
                    if (isShort && mappedData.salesRep) {
                        html += '<br><button type="button" class="action-btn-secondary" style="margin-top: 4px; background: #e67e22;" onclick="emailSalesRep(\'' + dataId + '\')" title="Email sales rep about short authorization">Email Sales Rep</button>';
                    }

                    // Create Note button for Sales Order - Secondary action (if notes enabled)
                    if (enableNotes && mappedData.salesOrderId) {
                        html += '<div id="' + noteDataId + '" class="hidden-data">';
                        html += '<input type="hidden" name="action" value="create_note">';
                        html += '<input type="hidden" name="transactionId" value="' + escapeHtml(mappedData.salesOrderId) + '">';
                        html += '</div>';
                        html += '<br><button type="button" class="action-btn-secondary" style="margin-top: 4px;" onclick="promptAndCreateNote(\'' + noteDataId + '\')" title="Add a note to this sales order">+ Note</button>';
                    }

                    html += '</td>';

                } else if (actionType === 'payment') {
                    var paymentData = extractRowData(resultsRange[i], actionType);
                    var dataId = 'payment_data_' + i;
                    var noteDataId = 'note_data_' + i;
                    
                    // Get Wells Fargo Authorization records from pre-loaded batch map
                    var wfAuthData = { records: [], total: 0, authNumbers: [] };
                    if (paymentData.salesOrderId && wfAuthMap[paymentData.salesOrderId]) {
                        wfAuthData = wfAuthMap[paymentData.salesOrderId];
                    }

                    html += '<td class="action-cell">';

                    // Check if this is a Credit Memo
                    if (paymentData.transactionType === 'Credit Memo') {
                        html += '<span style="color: #666; font-style: italic; font-size: 11px;">Refund Manually</span>';
                    } else {
                        // Hidden data container for payment (NO FORM TAG - just a div with data)
                        html += '<div id="' + dataId + '" class="hidden-data">';
                        html += '<input type="hidden" name="action" value="create_payment">';
                        html += '<input type="hidden" name="customer" value="' + escapeHtml(paymentData.customerId) + '">';
                        html += '<input type="hidden" name="amount" value="' + escapeHtml(paymentData.amount) + '">';
                        html += '<input type="hidden" name="wfAuthNumber" value="' + escapeHtml(paymentData.wfAuthNumber) + '">';
                        html += '<input type="hidden" name="invoiceNumber" value="' + escapeHtml(paymentData.invoiceNumber) + '">';
                        html += '</div>';

                        // Primary action: Create Payment button
                        html += '<button type="button" class="action-btn" onclick="promptAndSubmitPayment(\'' + dataId + '\', \'' + escapeHtml(paymentData.amount) + '\')">Create Payment</button>';
                    }

                    // Create Note button for ALL rows (invoices and credit memos) - Secondary action
                    // Only show if we have a valid transaction ID
                    if (paymentData.invoiceId) {
                        html += '<div id="' + noteDataId + '" class="hidden-data">';
                        html += '<input type="hidden" name="action" value="create_note">';
                        html += '<input type="hidden" name="transactionId" value="' + escapeHtml(paymentData.invoiceId) + '">';
                        html += '</div>';
                        html += '<br><button type="button" class="action-btn-secondary" style="margin-top: 4px;" onclick="promptAndCreateNote(\'' + noteDataId + '\')" title="Add a note to this transaction">+ Note</button>';
                    }

                    html += '</td>';
                }

                // Add regular data columns with selective HTML rendering
                // For deposit action type, we need to insert Deposit Validation column
                var colsToDisplay = (actionType === 'deposit') ? expectedColumns - 1 : expectedColumns;
                
                for (var col = 0; col < colsToDisplay; col++) {
                    // For deposit rows, insert Deposit Validation column after column 7 (Selling Location)
                    if (actionType === 'deposit' && col === 8) {
                        // Build Deposit Validation display
                        var validationHtml = '<div class="deposit-validation">';
                        
                        var soTotal = parseFloat(mappedData.salesOrderTotal) || 0;
                        var required = soTotal / 2;
                        var priorDeposits = parseFloat(mappedData.customerDepositTotal) || 0;
                        var wfCharge = parseFloat(mappedData.amount) || 0;
                        var afterProcessing = priorDeposits + wfCharge;
                        var variance = afterProcessing - required;
                        
                        validationHtml += 'SO Total: ' + soTotal.toFixed(2) + '\n\n';
                        validationHtml += 'Required (50% CD): ' + required.toFixed(2) + '\n\n';
                        validationHtml += 'Prior CDs: ' + priorDeposits.toFixed(2) + '\n\n';
                        validationHtml += 'WF To Be Processed: ' + wfCharge.toFixed(2) + '\n\n';
                        validationHtml += 'Total CDs After Processing: ' + afterProcessing.toFixed(2) + '\n\n';
                        
                        // Add validation indicator with color coding
                        if (variance === 0) {
                            validationHtml += '<span class="validation-match">✓ VALIDATED</span>';
                        } else if (variance < 0) {
                            validationHtml += '<span class="validation-short">⚠️ SHORT ' + Math.abs(variance).toFixed(2) + '</span>';
                        } else {
                            validationHtml += '<span class="validation-over">⚠️ OVER ' + variance.toFixed(2) + '</span>';
                        }
                        
                        validationHtml += '</div>';
                        html += '<td>' + validationHtml + '</td>';
                    }
                    
                    try {
                        var cellValue = '';
                        if (resultsRange[i].columns[col]) {
                            cellValue = resultsRange[i].getValue(resultsRange[i].columns[col]) || '';
                            var textValue = resultsRange[i].getText(resultsRange[i].columns[col]);
                            if (textValue && textValue !== 'undefined' && textValue !== cellValue) {
                                cellValue = textValue;
                            }
                        }

                        // Check if this column should allow HTML rendering
                        var columnLabel = resultsRange[i].columns[col] ?
                            (resultsRange[i].columns[col].label || '') : '';

                        // Make specific columns clickable based on column label
                        if (actionType === 'deposit') {
                            // For deposit table: Sales Order #, Customer, Wells Fargo Authorization Record ID
                            if (columnLabel === 'Sales Order #' && mappedData.salesOrderId) {
                                var soUrl = url.resolveRecord({
                                    recordType: record.Type.SALES_ORDER,
                                    recordId: mappedData.salesOrderId,
                                    isEditMode: false
                                });
                                html += '<td><a href="' + escapeHtml(soUrl) + '" target="_blank">' + escapeHtml(String(cellValue)) + '</a></td>';
                            } else if (columnLabel === 'Customer' && mappedData.customerId) {
                                var custUrl = url.resolveRecord({
                                    recordType: record.Type.CUSTOMER,
                                    recordId: mappedData.customerId,
                                    isEditMode: false
                                });
                                html += '<td><a href="' + escapeHtml(custUrl) + '" target="_blank">' + escapeHtml(String(cellValue)) + '</a></td>';
                            } else if (columnLabel === 'Wells Fargo Authorization Record ID' && mappedData.wfAuthId) {
                                var wfUrl = url.resolveRecord({
                                    recordType: 'customrecord_bas_wf_auth',
                                    recordId: mappedData.wfAuthId,
                                    isEditMode: false
                                });
                                html += '<td><a href="' + escapeHtml(wfUrl) + '" target="_blank">' + escapeHtml(String(cellValue)) + '</a></td>';
                            } else if (shouldAllowHtmlRendering(columnLabel)) {
                                html += '<td>' + sanitizeAllowedHtml(String(cellValue)) + '</td>';
                            } else {
                                html += '<td>' + escapeHtml(String(cellValue)) + '</td>';
                            }
                        } else if (actionType === 'payment') {
                            // For payment table: Document #, Customer, WF Record ID
                            if ((columnLabel === 'Document #' || columnLabel === 'Document Number') && paymentData.invoiceId) {
                                var recordType = (paymentData.transactionType === 'Credit Memo') ? record.Type.CREDIT_MEMO : record.Type.INVOICE;
                                var docUrl = url.resolveRecord({
                                    recordType: recordType,
                                    recordId: paymentData.invoiceId,
                                    isEditMode: false
                                });
                                html += '<td><a href="' + escapeHtml(docUrl) + '" target="_blank">' + escapeHtml(String(cellValue)) + '</a></td>';
                            } else if (columnLabel === 'Customer' && paymentData.customerId) {
                                var custUrl = url.resolveRecord({
                                    recordType: record.Type.CUSTOMER,
                                    recordId: paymentData.customerId,
                                    isEditMode: false
                                });
                                html += '<td><a href="' + escapeHtml(custUrl) + '" target="_blank">' + escapeHtml(String(cellValue)) + '</a></td>';
                            } else if (columnLabel === 'WF Record ID') {
                                // Display all WF records with their names and amounts
                                if (wfAuthData.records.length > 0) {
                                    var wfRecordsHtml = '';
                                    for (var wfIdx = 0; wfIdx < wfAuthData.records.length; wfIdx++) {
                                        var wfRec = wfAuthData.records[wfIdx];
                                        var wfUrl = url.resolveRecord({
                                            recordType: 'customrecord_bas_wf_auth',
                                            recordId: wfRec.id,
                                            isEditMode: false
                                        });
                                        if (wfIdx > 0) {
                                            wfRecordsHtml += '<br>';
                                        }
                                        wfRecordsHtml += '<a href="' + escapeHtml(wfUrl) + '" target="_blank">' + 
                                                        escapeHtml(wfRec.name) + '</a> ($' + 
                                                        wfRec.amount.toFixed(2) + ')';
                                    }
                                    html += '<td>' + wfRecordsHtml + '</td>';
                                } else {
                                    html += '<td style="color: #999; font-style: italic;">No WF records</td>';
                                }
                            } else if (columnLabel === 'Wells Fargo Authorization Amount ($)' || 
                                       columnLabel === 'WF Auth Amount ($)' ||
                                       columnLabel === 'WF Authorization Amount' ||
                                       columnLabel === 'Authorization Amount') {
                                // Display the TOTAL of all WF authorization amounts
                                if (wfAuthData.total > 0) {
                                    html += '<td>$' + wfAuthData.total.toFixed(2) + '</td>';
                                } else {
                                    html += '<td>$0.00</td>';
                                }
                            } else if (columnLabel === 'Wells Fargo Authorization #' || 
                                       columnLabel === 'WF Authorization #' ||
                                       columnLabel === 'Authorization #') {
                                // Display all authorization numbers
                                if (wfAuthData.authNumbers.length > 0) {
                                    html += '<td>' + escapeHtml(wfAuthData.authNumbers.join(', ')) + '</td>';
                                } else {
                                    html += '<td style="color: #999; font-style: italic;">-</td>';
                                }
                            } else if (columnLabel === 'Internal ID' || columnLabel === 'Transaction Internal ID' || columnLabel === 'Invoice Internal ID') {
                                // Make Internal ID clickable for invoices/credit memos
                                if (cellValue) {
                                    var recordTypeToUse = (paymentData.transactionType === 'Credit Memo') ? record.Type.CREDIT_MEMO : record.Type.INVOICE;
                                    var transactionUrl = url.resolveRecord({
                                        recordType: recordTypeToUse,
                                        recordId: cellValue,
                                        isEditMode: false
                                    });
                                    html += '<td><a href="' + escapeHtml(transactionUrl) + '" target="_blank">' + escapeHtml(String(cellValue)) + '</a></td>';
                                } else {
                                    html += '<td>' + escapeHtml(String(cellValue)) + '</td>';
                                }
                            } else if (shouldAllowHtmlRendering(columnLabel)) {
                                html += '<td>' + sanitizeAllowedHtml(String(cellValue)) + '</td>';
                            } else {
                                html += '<td>' + escapeHtml(String(cellValue)) + '</td>';
                            }
                        } else {
                            if (shouldAllowHtmlRendering(columnLabel)) {
                                html += '<td>' + sanitizeAllowedHtml(String(cellValue)) + '</td>';
                            } else {
                                html += '<td>' + escapeHtml(String(cellValue)) + '</td>';
                            }
                        }
                    } catch (e) {
                        html += '<td>Error</td>';
                    }
                }

                // Add User Notes column if notes are enabled
                if (enableNotes) {
                    if (recentNote && recentNote.note) {
                        // Format: "MM/DD/YYYY - Author Name - Note text here [EDIT]"
                        var noteDisplay = '';
                        if (recentNote.notedate) {
                            noteDisplay += escapeHtml(recentNote.notedate) + ' - ';
                        }
                        if (recentNote.author) {
                            noteDisplay += escapeHtml(recentNote.author) + ' - ';
                        }
                        noteDisplay += escapeHtml(recentNote.note);
                        
                        // Add [EDIT] link if we have a note ID
                        if (recentNote.id) {
                            var noteUrl = url.resolveRecord({
                                recordType: record.Type.NOTE,
                                recordId: recentNote.id,
                                isEditMode: true
                            });
                            noteDisplay += ' <a href="' + escapeHtml(noteUrl) + '" target="_blank" style="font-weight: bold; color: #4CAF50;">[EDIT]</a>';
                        }

                        html += '<td class="note-cell">' + noteDisplay + '</td>';
                    } else {
                        html += '<td class="note-cell" style="color: #999; font-style: italic;">No notes</td>';
                    }
                }

                html += '</tr>';
            }
            html += '</tbody></table>';

            return html;

        } catch (e) {
            log.error('Error building table for search ' + searchId, e.message);
            return '<div class="error-msg">Error loading search ' + escapeHtml(searchId) + ': ' + escapeHtml(e.message) + '</div>';
        }
    }

    /**
     * Extracts relevant data from a search result row for pre-populating forms
     * @param {Object} result - Search result row
     * @param {string} actionType - Type of action being performed
     * @returns {Object} Extracted data object
     */
    function extractRowData(result, actionType) {
        var data = {};

        try {
            if (actionType === 'payment') {
                // Extract column data including invoice number and transaction type
                var invoiceNumber = '';
                var transactionType = '';

                // Extract data from columns
                for (var i = 0; i < result.columns.length; i++) {
                    var column = result.columns[i];
                    var label = column.label || '';
                    var value = result.getValue(column) || '';

                    switch (label) {
                        case 'Document #':
                        case 'Document Number':
                        case 'Invoice':
                            invoiceNumber = value;
                            data.invoiceNumber = value;
                            break;
                        case 'Type':
                            transactionType = result.getText(column) || value;
                            data.transactionType = transactionType;
                            break;
                        case 'Customer':
                        case 'Customer Internal ID':
                            data.customerId = value;
                            break;
                        case 'Amount Remaining':
                        case 'Amount':
                            data.amount = value;
                            break;
                        case 'Wells Fargo Authorization #':
                        case 'WF Auth #':
                        case 'WF Authorization #':
                            data.wfAuthNumber = value;
                            break;
                        case 'Internal ID':
                        case 'Transaction Internal ID':
                        case 'Invoice Internal ID':
                            data.invoiceId = value;
                            break;
                        case 'Sales Order Internal ID':
                            data.salesOrderId = value;
                            break;
                        default:
                            break;
                    }
                }
                
                // Fallback to result.id if no internal ID found in columns
                if (!data.invoiceId) {
                    data.invoiceId = result.id;
                }
            } else if (actionType === 'deposit') {
                // Keep existing deposit logic using mapDepositColumns
                var mappedData = mapDepositColumns(result);
                data.customerId = mappedData.customerId;
                data.salesOrderId = mappedData.salesOrderId;
                data.amount = mappedData.amount;
                data.wfAuthId = mappedData.wfAuthId;
                data.location = mappedData.location;
                data.wfAuthNumber = mappedData.wfAuthNumber;
            }
        } catch (e) {
            log.error('Error extracting row data', {
                error: e.message,
                stack: e.stack
            });
        }

        return data;
    }
    /**
     * Determines if a column should allow HTML rendering based on column label
     * @param {string} columnLabel - The column label to check
     * @returns {boolean} True if HTML should be allowed
     */
    function shouldAllowHtmlRendering(columnLabel) {
        var htmlColumns = ['Terms Summary', 'Manufacturers'];
        return htmlColumns.indexOf(columnLabel) !== -1;
    }

    /**
     * Sanitizes HTML content to allow only safe tags
     * @param {string} html - HTML content to sanitize
     * @returns {string} Sanitized HTML content
     */
    function sanitizeAllowedHtml(html) {
        if (!html) return '';

        // Only allow specific safe HTML tags
        var allowedTags = {
            'b': true,
            'strong': true,
            'br': true,
            'i': true,
            'em': true
        };

        // Remove any script tags and event handlers first
        var cleaned = html.toString()
            .replace(/<script[^>]*>.*?<\/script>/gi, '')
            .replace(/on\w+\s*=\s*["'][^"']*["']/gi, '')
            .replace(/javascript:/gi, '');

        // Simple tag validation - only allow whitelisted tags
        cleaned = cleaned.replace(/<(\/?)([\w]+)([^>]*)>/gi, function (match, slash, tagName, attributes) {
            var lowerTagName = tagName.toLowerCase();
            if (allowedTags[lowerTagName]) {
                // For allowed tags, remove any attributes (for simplicity)
                if (lowerTagName === 'br') {
                    return '<' + slash + lowerTagName + '>';
                } else {
                    return '<' + slash + lowerTagName + '>';
                }
            }
            return ''; // Remove disallowed tags
        });

        return cleaned;
    }

    /**
     * Enhanced HTML escape function with comprehensive character coverage
     * @param {string} text - Text to escape
     * @returns {string} Escaped text
     */
    function escapeHtml(text) {
        if (!text) return '';
        return text.toString()
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;')
            .replace(/\//g, '&#x2F;');
    }

    /**
     * Maps Wells Fargo deposit search columns by header name to extract relevant data
     * @param {Object} result - Search result row
     * @returns {Object} Mapped deposit data
     */
    function mapDepositColumns(result) {
        var mappedData = {
            customerId: '',
            salesOrderId: '',
            amount: '',
            wfAuthId: result.id,
            location: '',
            wfAuthNumber: '',
            salesOrderTotal: '',
            customerDepositTotal: '',
            salesRep: ''
        };

        try {
            for (var i = 0; i < result.columns.length; i++) {
                var column = result.columns[i];
                var label = column.label || '';
                var value = result.getValue(column) || '';

                switch (label) {
                    case 'Customer Internal ID':
                        mappedData.customerId = value;
                        break;
                    case 'Sales Order Internal ID':
                        mappedData.salesOrderId = value;
                        break;
                    case 'Customer Deposit Amount':
                        mappedData.amount = value;
                        break;
                    case 'Selling Location':
                        mappedData.location = value; // This is the department ID
                        break;
                    case 'Wells Fargo Authorization #':
                    case 'WF Auth #':
                        mappedData.wfAuthNumber = value;
                        break;
                    case 'Sales Rep':
                        mappedData.salesRep = value;
                        break;
                    default:
                        break;
                }
            }

            // Get Sales Order Total if we have a Sales Order ID
            if (mappedData.salesOrderId) {
                var soTotal = getSalesOrderTotal(mappedData.salesOrderId);
                mappedData.salesOrderTotal = soTotal.toFixed(2);
                
                // Get Applied Customer Deposits Total
                var cdTotal = getAppliedCustomerDeposits(mappedData.salesOrderId);
                mappedData.customerDepositTotal = cdTotal.toFixed(2);
            }

        } catch (e) {
            log.error('Error mapping deposit columns', e.message);
        }

        return mappedData;
    }

    /**
     * Generates NetSuite URL for creating a record with pre-populated data
     * @param {string} recordType - The type of record to create
     * @param {Object} data - Data object containing field values
     * @returns {string} Generated URL string
     */
    function generateRecordUrl(recordType, data) {
        try {
            var baseUrl = url.resolveRecord({
                recordType: recordType,
                isEditMode: false
            });

            var params = [];
            for (var key in data) {
                if (data[key]) {
                    params.push(encodeURIComponent(key) + '=' + encodeURIComponent(data[key]));
                }
            }

            if (params.length > 0) {
                baseUrl += (baseUrl.indexOf('?') > -1 ? '&' : '?') + params.join('&');
            }

            return baseUrl;

        } catch (e) {
            log.error('Error generating record URL', e.message);
            return '#';
        }
    }

    /**
     * Escapes HTML characters to prevent XSS
     * @param {string} text - Text to escape
     * @returns {string} Escaped text
     */
    function escapeHtml(text) {
        if (!text) return '';
        return text.toString()
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    /**
    * Finds an invoice by document number for a specific customer
    * @param {number} customerId - The customer internal ID
    * @param {string} invoiceNumber - The invoice document number
    * @returns {number|null} The invoice internal ID or null if not found
    */
    function findInvoiceByNumber(customerId, invoiceNumber) {
        try {
            var invoiceSearch = search.create({
                type: search.Type.INVOICE,
                filters: [
                    ['entity', 'anyof', customerId],
                    'AND',
                    ['tranid', 'is', invoiceNumber],
                    'AND',
                    ['mainline', 'is', 'T']
                ],
                columns: ['internalid']
            });

            var searchResults = invoiceSearch.run().getRange({
                start: 0,
                end: 1
            });

            if (searchResults.length > 0) {
                return parseInt(searchResults[0].getValue('internalid'), 10);
            }

            return null;

        } catch (e) {
            log.error('Error finding invoice by number', {
                error: e.message,
                customerId: customerId,
                invoiceNumber: invoiceNumber
            });
            return null;
        }
    }

    return {
        onRequest: onRequest
    };
});