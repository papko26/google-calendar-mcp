import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { OAuth2Client } from "google-auth-library";
import { BaseToolHandler } from "./BaseToolHandler.js";
import { calendar_v3 } from 'googleapis';
import { buildListFieldMask } from "../../utils/field-mask-builder.js";
import { createStructuredResponse } from "../../utils/response-builder.js";
import { SearchEventsResponse, StructuredEvent, convertGoogleEventToStructured, ExtendedEvent } from "../../types/structured-responses.js";

interface SearchEventsFlexibleArgs {
    calendarId: string | string[];
    pattern: string;
    useRegex?: boolean;
    caseSensitive?: boolean;
    timeMin?: string;
    timeMax?: string;
    timeZone?: string;
    fields?: string[];
    account?: string | string[];
}

export class SearchEventsFlexibleHandler extends BaseToolHandler {
    async runTool(args: SearchEventsFlexibleArgs, accounts: Map<string, OAuth2Client>): Promise<CallToolResult> {
        const calendarNamesOrIds = Array.isArray(args.calendarId)
            ? args.calendarId
            : [args.calendarId];

        const selectedAccounts = this.getClientsForAccounts(args.account, accounts);
        const resolutionWarnings: string[] = [];

        let accountCalendarMap: Map<string, string[]>;

        if (selectedAccounts.size > 1 || calendarNamesOrIds.length > 1) {
            const { resolved, warnings } = await this.calendarRegistry.resolveCalendarsToAccounts(
                calendarNamesOrIds,
                selectedAccounts
            );
            accountCalendarMap = resolved;
            resolutionWarnings.push(...warnings);

            if (accountCalendarMap.size === 0) {
                await this.throwNoCalendarsFoundError(calendarNamesOrIds, selectedAccounts);
            }
        } else {
            const { accountId, calendarId } = await this.getClientWithAutoSelection(
                args.account,
                calendarNamesOrIds[0],
                accounts,
                'read'
            );
            accountCalendarMap = new Map([[accountId, [calendarId]]]);
        }

        // Build matcher function
        const matcher = this.buildMatcher(args.pattern, args.useRegex ?? false, args.caseSensitive ?? false);

        const allEvents: ExtendedEvent[] = [];
        const queriedCalendarIds: string[] = [];

        await Promise.all(
            Array.from(accountCalendarMap.entries()).map(async ([accountId, calendarIds]) => {
                const client = selectedAccounts.get(accountId)!;
                for (const calendarId of calendarIds) {
                    try {
                        const events = await this.listAllEvents(client, calendarId, args);
                        for (const event of events) {
                            if (this.eventMatchesPattern(event, matcher)) {
                                allEvents.push({ ...event, calendarId, accountId });
                            }
                        }
                        queriedCalendarIds.push(calendarId);
                    } catch (error) {
                        if (accountCalendarMap.size > 1 || calendarIds.length > 1) {
                            const message = error instanceof Error ? error.message : String(error);
                            resolutionWarnings.push(`Failed to search calendar "${calendarId}": ${message}`);
                        } else {
                            throw error;
                        }
                    }
                }
            })
        );

        this.sortEventsByStartTime(allEvents);

        const structuredEvents: StructuredEvent[] = allEvents.map(event =>
            convertGoogleEventToStructured(event, event.calendarId, event.accountId)
        );

        const response: SearchEventsResponse = {
            events: structuredEvents,
            totalCount: allEvents.length,
            query: args.pattern,
            ...(queriedCalendarIds.length === 1 && { calendarId: queriedCalendarIds[0] }),
            ...(queriedCalendarIds.length > 1 && { calendars: queriedCalendarIds }),
            ...(selectedAccounts.size > 1 && { accounts: Array.from(selectedAccounts.keys()) }),
            ...(resolutionWarnings.length > 0 && { warnings: resolutionWarnings })
        };

        if (args.timeMin || args.timeMax) {
            const firstAccountId = accountCalendarMap.keys().next().value as string;
            const firstCalendarId = accountCalendarMap.get(firstAccountId)?.[0] || 'primary';
            const client = selectedAccounts.get(firstAccountId)!;
            const { timeMin, timeMax } = await this.normalizeTimeRange(
                client, firstCalendarId, args.timeMin, args.timeMax, args.timeZone
            );
            response.timeRange = { start: timeMin || '', end: timeMax || '' };
        }

        return createStructuredResponse(response);
    }

    private buildMatcher(pattern: string, useRegex: boolean, caseSensitive: boolean): (text: string) => boolean {
        if (useRegex) {
            const flags = caseSensitive ? '' : 'i';
            const re = new RegExp(pattern, flags);
            return (text) => re.test(text);
        }
        if (!caseSensitive) {
            const lower = pattern.toLowerCase();
            return (text) => text.toLowerCase().includes(lower);
        }
        return (text) => text.includes(pattern);
    }

    private eventMatchesPattern(event: calendar_v3.Schema$Event, matcher: (text: string) => boolean): boolean {
        return (
            matcher(event.summary ?? '') ||
            matcher(event.description ?? '') ||
            matcher(event.location ?? '')
        );
    }

    private async listAllEvents(
        client: OAuth2Client,
        calendarId: string,
        args: SearchEventsFlexibleArgs
    ): Promise<calendar_v3.Schema$Event[]> {
        const calendar = this.getCalendar(client);
        const { timeMin, timeMax } = await this.normalizeTimeRange(
            client, calendarId, args.timeMin, args.timeMax, args.timeZone
        );

        const fieldMask = buildListFieldMask(args.fields);
        const allItems: calendar_v3.Schema$Event[] = [];
        let pageToken: string | undefined;

        do {
            const response = await calendar.events.list({
                calendarId,
                timeMin,
                timeMax,
                singleEvents: true,
                orderBy: 'startTime',
                maxResults: 2500,
                pageToken,
                ...(fieldMask && { fields: fieldMask }),
            });
            allItems.push(...(response.data.items ?? []));
            pageToken = response.data.nextPageToken ?? undefined;
        } while (pageToken);

        return allItems;
    }
}
