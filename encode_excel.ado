// Author: Sally Hudson
// Created: October 2012
// Modified: April 2014
// Purpose: This program maps raw values to clean, labeled integers using a crosswalk 
// provided in an external spreadsheet.


capture program drop encode_excel

program define encode_excel, nclass
	syntax varlist using/, raw(string) clean(string) label(string) [sheet(string)] [missing(string)]

	preserve
	di "encoding `varlist' from `using'..."
	
	// declare temporary variables
	tempvar merge code
	
	// get code values matched to raw values 
	import excel  `using', sheet(`sheet') firstrow clear
	keep `clean' `raw' `label'
	
		// allow for raw and label to be same spreadsheet column
		if "`label'" == "`raw'" {
			tempvar label
			gen `label' = `raw'
		}
		// allow for raw and clean to be same spreadsheet column
		if "`clean'" == "`raw'" {
			tempvar clean
			gen `clean' = `raw'
		}
		
	// determine if merge variable is string or numeric
	qui cap destring `raw', replace
	local type_string = 0
	foreach v of varlist `raw' {
		cap confirm string variable `v'
		if _rc == 0 {
			local type_string = 1
		}
	}

	// reshape raw values if more than one supplied
	rename (`raw') merge#, addnumber
	cap confirm variable merge2
	if _rc == 0 {
		gen i = _n
		qui tostring merge*, replace
		qui reshape long merge, i(i) j(j)
		gen `merge' = merge
		drop i j merge*
	}
	else {
		gen `merge' = merge1
		drop merge*
	}

	// verify that only one label is supplied for each clean code value	
	bysort `label': egen clean_max = max(`clean')
	bysort `label': egen clean_min = min(`clean')
	qui count if clean_max != clean_min & !missing(`label')
	if r(N) != 0 {
		display as error "The following code values are assigned to more than one label"
		list `clean' `label' if clean_max != clean_min
		STOP
	}

	// define label : this part modified by Sookyo Jeong in June 2015
	//					sookyojeong@gmail.com
	tempfile beforelabel
	save `beforelabel', replace
	
	bys `clean': gen count=_n
	keep if count==1
	qui levelsof `clean', local(codes_clean) 
	foreach x of local codes_clean{
	forvalues i = 1/`=_N'{
		if `clean'[`i'] == `x' {
				local label_`x' = `label'[`i']
				break
				}
		}
	}
	use `beforelabel', clear

	// reduce data set to unique combinations of raw, clean, and label values 
	qui drop if missing(`merge')
	cap confirm string variable `merge' 
	if !_rc {
		qui drop if `merge' == "."
	}
	qui duplicates drop
		
	// save matched codes data set	
	gen `code' = `clean'
	tempfile codes
	qui save `codes', replace

	// encode and label variables
	restore



	
	foreach v of varlist `varlist' {

		// merge with clean values
		rename `v' `merge'
		if `type_string' {
			qui tostring `merge', replace
			replace `merge' = "" if `merge' == "."
		}
		else {
			qui destring `merge', replace
		}
		qui merge m:1 `merge' using `codes', keep(master match) keepusing(`code')
		
		// verify that all raw values have found a match
		qui count if !missing(`merge') & missing(`code')
		if r(N) > 0 {
			display as error "The following values for `v' could not be encoded:"
			qui gen uncoded = `merge' if !missing(`merge') & missing(`code')
			tab `merge' if !missing(`merge') & missing(`code')
			STOP
		}
		
		// label variable
		local varlabel: variable label `merge'
		qui gen `v' = `code'
		label variable `merge' "`varlabel'"
		
		// label values
		cap label drop `v'
		local i = 1
		foreach x of local codes_clean {
			local notMissing = 1
			if "`missing'" != "" {
				foreach l_missing in `missing' {
					if ("`label_`x''" == "`l_missing'") {
						local notMissing = 0
					}
				}
			}
			if `notMissing' {
				label define `v' `x' "`label_`x''", modify
			}
			else {
				replace `v' = . if `v' == `x'
			}
		}
		label values `v' `v'
		drop `merge' `code' _merge
	di "here4"
	
		// tab new codes
		qui compress `v'
		tab `v', missing
		di ""
	}
	end	

