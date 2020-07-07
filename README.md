# VBA-challenge

### Script Files

* [Basic Assignment with Both Challenges](Resources/Module_full_assignment.bas) - the basic assignment and both challenges. **Use this to run the entire assignment.**
* [Basic Assignment](Resources/Module_basic.bas) - just the basic assignment
* [Basic Assignment and Challenge 1](Resources/Module_basic_and_challenge1.bas) - the basic assignment and the 1st challenge


### Screenshot Files

* [2014](Resources/2014_Ticker_Screenshot.PNG) - Results for 2014
* [2015](Resources/2015_Ticker_Screenshot.PNG) - Results for 2015
* [2016](Resources/2016_Ticker_Screenshot.PNG) - Results for 2016

## Notes

Not sure if I needed to include all the scripts but they were built as I was figuring out each part. You should just have to use Module_full_assignment script (the Basic Assignment with Both Challenges link above) to accomplish all parts of this assignment. No real need for the other two scripts other than to documnent my thought process more or less.

The scripts were modified along the way to getting the one with both challenges because I found a different way to do things or the original didn't work with the additions to get a part of the challenge to work. I wasn't sure if sorting needed to be a part of this but just in case, I added that. Initially that was done as a separate subroutine (SortWorksheet) which was then called from the ticker_totals subroutine. But I found it worked better as just part of the main subroutine by the time I was working on the final challenge so that SortWorksheet subroutine is not separate in Module_full_assignment script.

I had some trouble looping through the worksheets and getting the 3 Challenge totals to work correctly. It kept using the totals from all 3 worksheets as it moved from worksheet to worksheet. So I found that if I set the variables holding the final totals to 0 first, it would then use just that current worksheet's data. I'm sure there were other ways to accomplish this but this is how I got it to work.

All the scripts and screenshots should be in the Resources folder in case the links don't work in this README.md.