# Configure Update Rings

The Update Ring feature in the EAM-AutoUpdater allows you to delay the availability of a new software package for certain assignment groups. This is useful for business-critical applications that require additional validation before being rolled out to a wider audience.

When a matching update ring entry is found for an assignment, the script sets the `installTimeSettings.startDateTime` on the Intune assignment so the app only becomes available after the configured delay.

## Append Parameter

In order to use the Update Ring feature, append the `-UpdateRings` switch to your `Invoke-EAMAutoupdate` call:

```powershell
Invoke-EAMAutoupdate -TeamsWebhookUri <URI> -UpdateESP -UpdateRings
```

## Configure the `$UpdateRingSettings` Variable

The `$UpdateRingSettings` variable is an array of `PSCustomObject` entries defined before the `Invoke-EAMAutoupdate` call. Each entry targets a specific app, assignment type, and group combination.

### Properties

| Property | Required | Description |
|---|---|---|
| `ApplicationName` | Yes | The display name of the EAM catalog app (must match exactly). |
| `Assignmenttype` | Yes | The assignment intent to match: `available`, `required`, or `uninstall`. |
| `groupId` | Yes | The Entra ID object ID of the target group, or `adadadad-808e-44e2-905a-0b7873a8a531` for **All Devices** / `acacacac-9df4-4c7d-9d50-4ef0226f57a9` for **All Users**. |
| `DaysinDelay` | Yes | Number of days to delay availability from the moment the script runs. |
| `AvailabilityTimeHour` | No | The hour of the day (0–23) at which the app should become available. If omitted, the time component is midnight (00:00). |
| `TimeZoneId` | No | A Windows time zone ID (e.g. `W. Europe Standard Time`) that determines how `DaysinDelay` and `AvailabilityTimeHour` are interpreted. The behaviour depends on the assignment intent — see [Time Zone Calculations](#time-zone-calculations) for details. If omitted, UTC is used. |

> **Tip:** To list all available time zone IDs on your system, run:
> ```powershell
> [System.TimeZoneInfo]::GetSystemTimeZones() | Select-Object Id, DisplayName
> ```
> You can also look up a specific region by filtering the output:
> ```powershell
> [System.TimeZoneInfo]::GetSystemTimeZones() | Where-Object DisplayName -like '*Europe*'
> ```
> Use the value from the **Id** column as the `TimeZoneId` property.

### Example
The following example is contains more information: 

```powershell
$UpdateRingSettings = @(
    # Ring 1 – Pilot group: available after 3 days at 09:00 W. Europe Standard Time
    [PSCustomObject]@{
        ApplicationName = 'Chrome for Business 64-bit'
        Assignmenttype  = 'available'
        groupId         = '223f4e5b-61a6-47cc-a13c-e2649bc3ad31'
        DaysinDelay     = '3'
        AvailabilityTimeHour = 9
        TimeZoneId      = 'W. Europe Standard Time'
    }
    # Ring 2 – All Devices: available after 7 days at midnight UTC (default)
    [PSCustomObject]@{
        ApplicationName = 'Chrome for Business 64-bit'
        Assignmenttype  = 'available'
        groupId         = 'adadadad-808e-44e2-905a-0b7873a8a531'
        DaysinDelay     = '7'
    }
)
```

Additionally you can also format the object like this: 

```powershell
$UpdateRingSettings = @(
    [PSCustomObject]@{ApplicationName = 'Chrome for Business 64-bit'; Assignmenttype = 'available'; groupId = '223f4e5b-61a6-47cc-a13c-e2649bc3ad31'; DaysinDelay = '3'; AvailabilityTimeHour = 9; TimeZoneId = 'W. Europe Standard Time' }
    [PSCustomObject]@{ApplicationName = 'Chrome for Business 64-bit'; Assignmenttype = 'available'; groupId = 'adadadad-808e-44e2-905a-0b7873a8a531'; DaysinDelay = '7' }
)
```

#### Time Zone Calculations

The script handles the `startDateTime` and `useLocalTime` properties on Intune's `installTimeSettings` differently depending on the assignment intent and whether `TimeZoneId` is specified.

> **Important:** Intune only supports `useLocalTime = $true` for **required** assignments. For **available** assignments, `useLocalTime` must be `$false` and the datetime must be in UTC. The script handles this automatically.

### When `TimeZoneId` is omitted (UTC mode)

The datetime is calculated in UTC and sent with `useLocalTime = $false`. Intune interprets and displays the value as UTC. This applies to all assignment intents.

1. Take the current UTC date (midnight, time component stripped).
2. Add the number of `DaysinDelay`.
3. If `AvailabilityTimeHour` is provided, add that many hours.
4. The result is sent as UTC.

**Example:** The script runs on `2026-04-11` at any time. `DaysinDelay = 7`, `AvailabilityTimeHour` is not set.
- Calculated start: `2026-04-18T00:00:00Z`
- Intune shows: **00:00 UTC** on April 18.

**Example:** Same date, `DaysinDelay = 7`, `AvailabilityTimeHour = 14`.
- Calculated start: `2026-04-18T14:00:00Z`
- Intune shows: **14:00 UTC** on April 18.

### When `TimeZoneId` is specified

The scheduled datetime is always calculated in the target time zone first:

1. Convert the current UTC time to the target time zone to determine "today" in that zone.
2. Take that local date (midnight) and add `DaysinDelay`.
3. If `AvailabilityTimeHour` is provided, add that many hours.

What happens next depends on the assignment intent:

#### Required assignments (`useLocalTime = $true`)

The local datetime is sent to Intune with the timezone's UTC offset appended (e.g. `2026-04-15T09:00:00+02:00`) and `useLocalTime = $true`. Intune interprets this as local device time — each device will install the app at 09:00 in its own local time zone.

**Example:** `DaysinDelay = 3`, `AvailabilityTimeHour = 9`, `TimeZoneId = 'W. Europe Standard Time'` (UTC+2 during CEST). Script runs on `2026-04-11 22:00 UTC`.
1. Current time in CEST: `2026-04-12 00:00` → local date is April 12.
2. Add 3 days → `2026-04-15`.
3. Add 9 hours → `2026-04-15 09:00`.
4. Sent to Intune: `2026-04-15T09:00:00+02:00` with `useLocalTime = $true`.
5. Intune shows: **09:00** on April 15 (device local time).

#### Available assignments (`useLocalTime = $false`)

Intune does not support `useLocalTime = $true` for available assignments. The script therefore converts the calculated local datetime back to UTC before sending it with `useLocalTime = $false`.

**Example:** Same configuration as above.
1. Local scheduled time: `2026-04-15 09:00 CEST`.
2. Convert to UTC: `2026-04-15 07:00:00Z` (CEST is UTC+2).
3. Sent to Intune: `2026-04-15T07:00:00Z` with `useLocalTime = $false`.
4. Intune shows: **07:00 UTC** on April 15 (which equals 09:00 CEST).

> **Note:** For available assignments, the Intune portal will display the UTC time. The actual availability on devices will still correspond to the intended local hour, but the portal representation is in UTC. The Teams notification shows the intended local time for clarity.

> **Note:** Because the date is first determined in the target time zone, running the script late in the UTC evening may result in a different "today" than expected. For example, `22:00 UTC` is already midnight in CEST, so the local date rolls forward.

#### Assignments Without a Matching Update Ring

When no matching entry exists in `$UpdateRingSettings` for a given assignment, or when the `-UpdateRings` switch is not used, no `installTimeSettings` are applied. The app becomes available as soon as possible. The Teams notification will show **As soon as possible** in the Availability column for these assignments.

#### Teams Notification

The adaptive card sent to Teams includes an **Availability** column in the assignment table. Each row shows:

![UpdateRingsNotification](./Screenshots/UpdateRingsNotification.png)

- **As soon as possible** — when no update ring delay was applied.
- **A date and time with the time zone** — when a `TimeZoneId` was configured (e.g. `2026-04-15 09:00 (W. Europe Standard Time)`).
- **A date and time in UTC** — when no `TimeZoneId` was specified (e.g. `2026-04-18 00:00 UTC`).

## Implement the Variable in the script

In order to make the configured variable usable in the script so that it can be picked up, you have to implement it before the `Invoke-EAMAutoUpdate` call. 

### Example 
```PowerShell
try {
    Connect-MgGraph -Identity
}
catch {
    throw "Failed to connect to Graph. Error: $_"
}

$UpdateRingSettings = @(
    [PSCustomObject]@{ApplicationName = 'Chrome for Business 64-bit'; Assignmenttype = 'available'; groupId = '223f4e5b-61a6-47cc-a13c-e2649bc3ad31'; DaysinDelay = '3'; AvailabilityTimeHour = 9; TimeZoneId = 'W. Europe Standard Time' }
    [PSCustomObject]@{ApplicationName = 'Chrome for Business 64-bit'; Assignmenttype = 'available'; groupId = 'adadadad-808e-44e2-905a-0b7873a8a531'; DaysinDelay = '7' }
)

Invoke-EAMAutoupdate -TeamsWebhookUri <"TeamsWebhookUri"> -UpdateESP -ExcludeApps <"draw.io Desktop"> -UpdateRings
```

