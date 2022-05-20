# Supply Line Utility

![Screenshot](screenshot.jpg)

## About

This utility automates updates to the Promise Date Verification using estimated completion dates provided by a scheduler. When updates cannot be made automatically or for exceptions, dates are left empty awaiting manual review.

## Installation
Download and run the latest release.

## Instructions for Use

1. Click the first 'browse' button, navigate to and select the spreadsheet file containing estimated dates from the scheduler.
2. Click the second 'browse' button, navigate to and select the spreadsheet file containing the Promise Date Verification template from the customer.
3. Enter a number of days to add to the estimated completion dates provided by the scheduler and select whether weekends should be included in the transit time calculation. Note that transit time does not include the estimated ship date.

Example: 
- Entering three (3) in 'Transit Time Days' and unchecking 'Include Weekends' will result in a shipment that is shipped on 5/6/2022 being promised for delivery on 5/11/2022.

## Issues or Feature Requests
Please submit issues [here](https://github.com/paulrunco/date-verification/issues) or email the author directly.