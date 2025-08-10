using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace OpportunityQuoteSumPlugin
{
    public class OpportunityQuoteSumPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Get the execution context
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                tracingService.Trace("Plugin started - Message: {0}, Entity: {1}", context.MessageName, context.PrimaryEntityName);

                // Check if this is an update operation on quote entity
                if (context.MessageName != "Update" || context.PrimaryEntityName != "quote")
                {
                    tracingService.Trace("Not a Quote Update operation, exiting");
                    return;
                }

                Guid quoteId = context.PrimaryEntityId;
                tracingService.Trace("Processing Quote ID: {0}", quoteId);

                // Retrieve the complete quote record after update
                Entity currentQuote = service.Retrieve("quote", quoteId, new ColumnSet("statuscode", "opportunityid", "rhs_totalamountrhs"));
                tracingService.Trace("Retrieved current quote record");

                // Check if statuscode is 4 (Won)
                if (!currentQuote.Contains("statuscode"))
                {
                    tracingService.Trace("StatusCode field not found on quote");
                    return;
                }

                OptionSetValue statusCode = currentQuote.GetAttributeValue<OptionSetValue>("statuscode");
                tracingService.Trace("Current StatusCode: {0}", statusCode?.Value);

                // Check if statuscode was actually changed in this update
                Entity target = null;
                bool statusWasChanged = false;

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    target = (Entity)context.InputParameters["Target"];
                    statusWasChanged = target.Contains("statuscode");
                }

                if (!statusWasChanged)
                {
                    tracingService.Trace("StatusCode was not changed in this update, no action needed");
                    return;
                }

                tracingService.Trace("StatusCode was changed - proceeding with recalculation");

                // Check if there's an opportunity reference
                if (!currentQuote.Contains("opportunityid"))
                {
                    tracingService.Trace("No opportunity reference found on this quote");
                    return;
                }

                EntityReference opportunityRef = currentQuote.GetAttributeValue<EntityReference>("opportunityid");
                if (opportunityRef == null)
                {
                    tracingService.Trace("Opportunity reference is null");
                    return;
                }

                tracingService.Trace("Processing opportunity: {0} - Recalculating total from ALL currently Won quotes", opportunityRef.Id);

                // Query all OTHER quotes related to this opportunity with statuscode = 4 (Won)
                // We'll handle the current quote separately to avoid timing issues
                QueryExpression quoteQuery = new QueryExpression("quote")
                {
                    ColumnSet = new ColumnSet("rhs_totalamountrhs", "statuscode", "statecode", "quoteid", "name"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };

                quoteQuery.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, opportunityRef.Id);
                quoteQuery.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 4);
                quoteQuery.Criteria.AddCondition("quoteid", ConditionOperator.NotEqual, quoteId); // Exclude current quote
                // Remove statecode condition - let's see if that's the issue
                // quoteQuery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 1);

                EntityCollection otherWonQuotes = service.RetrieveMultiple(quoteQuery);
                tracingService.Trace("Found {0} OTHER won quotes for opportunity {1} (excluding current quote)", otherWonQuotes.Entities.Count, opportunityRef.Id);

                // Calculate total from other won quotes
                decimal totalAmount = 0;
                int quotesWithAmount = 0;

                foreach (Entity quote in otherWonQuotes.Entities)
                {
                    tracingService.Trace("Processing other quote: {0}", quote.Id);

                    if (quote.Contains("rhs_totalamountrhs"))
                    {
                        Money amount = quote.GetAttributeValue<Money>("rhs_totalamountrhs");
                        if (amount != null && amount.Value > 0)
                        {
                            totalAmount += amount.Value;
                            quotesWithAmount++;
                            tracingService.Trace("Other Quote {0}: Added amount {1}, Running total: {2}", quote.Id, amount.Value, totalAmount);
                        }
                    }
                }

                // Now handle the current quote - add its amount if it's now Won (status = 4)
                if (statusCode != null && statusCode.Value == 4)
                {
                    if (currentQuote.Contains("rhs_totalamountrhs"))
                    {
                        Money currentAmount = currentQuote.GetAttributeValue<Money>("rhs_totalamountrhs");
                        if (currentAmount != null && currentAmount.Value > 0)
                        {
                            totalAmount += currentAmount.Value;
                            quotesWithAmount++;
                            tracingService.Trace("Current Quote {0}: Added amount {1}, Final total: {2}", quoteId, currentAmount.Value, totalAmount);
                        }
                        else
                        {
                            tracingService.Trace("Current Quote {0}: Amount is null or zero", quoteId);
                        }
                    }
                    else
                    {
                        tracingService.Trace("Current Quote {0}: rhs_totalamountrhs field not found", quoteId);
                    }
                }
                else
                {
                    tracingService.Trace("Current Quote {0}: Status is not Won ({1}), not including in total", quoteId, statusCode?.Value);
                }

                tracingService.Trace("Recalculated total for opportunity {0}: {1} from {2} Won quotes", opportunityRef.Id, totalAmount, quotesWithAmount);

                // Update the opportunity with the calculated total
                Entity opportunityUpdate = new Entity("opportunity", opportunityRef.Id);
                opportunityUpdate["new_qoutesamountcurrency"] = new Money(totalAmount);

                service.Update(opportunityUpdate);
                tracingService.Trace("Successfully updated opportunity {0} with amount {1}", opportunityRef.Id, totalAmount);
            }
            catch (Exception ex)
            {
                tracingService.Trace("ERROR: {0}", ex.Message);
                tracingService.Trace("STACK TRACE: {0}", ex.StackTrace);

                if (ex.InnerException != null)
                {
                    tracingService.Trace("INNER EXCEPTION: {0}", ex.InnerException.Message);
                }

                throw new InvalidPluginExecutionException($"QuoteStatusUpdatePlugin failed: {ex.Message}", ex);
            }
        }
    }
}

/* REGISTRATION INSTRUCTIONS:
 * 
 * 1. Message: Update
 * 2. Primary Entity: quote  
 * 3. Stage: Post-operation (40)
 * 4. Mode: Synchronous
 * 5. Filtering Attributes: statuscode (only trigger when status changes)
 * 6. NO post-images needed
 * 
 * HOW IT WORKS:
 * - Triggers whenever ANY quote status changes (not just to Won)
 * - Recalculates the total from ALL currently Won quotes for the opportunity
 * - This handles both: quotes becoming Won AND quotes changing from Won to other statuses
 * 
 * TESTING:
 * 1. Create quotes linked to an opportunity with values in rhs_totalamountrhs
 * 2. Change quote status TO "Won" → field should increase
 * 3. Change quote status FROM "Won" to something else → field should decrease
 * 4. Check Plugin Trace Log for execution details
 */