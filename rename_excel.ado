// Author: Sally Hudson
// Created: August 2012
// Modified: January 2014
// Purpose: This program renames and labels variables using an external Excel spreadsheet.

capture program drop rename_excel

program define rename_excel, nclass
	syntax using/, name_old(string) name_new(string) [sheet(string)] [label(string)] [cond_var(string)] [cond_value(string)] [dropx]
	** random changes
	// display comment
	display ""
	display "renaming variables from workbook `using'"
	display ""
	
	// preserve the existing data
	preserve
	
	// get variable names
	import excel `using', firstrow allstring sheet(`sheet') clear
	quietly keep if (!missing(`name_old'))
	
		// The cond_var option allows you to specify an additional column in the 
		// spreadsheet that identifies the relevant variables.
		if "`cond_var'" != "" & "`cond_value'" != "" {
			capture confirm numeric variable `cond_var' 
			if !_rc {
				keep if `cond_var' == `cond_value'
			}
			else {
				keep if `cond_var' == "`cond_value'"
			}
		
		}
	
	// verify that old variable names are unique
	qui duplicates tag `name_old', gen(dup)
	qui count if dup 
	if r(N) > 0 {
		display as error "The name_old column contains duplicate variable names.  Please specify a 1:1 mapping from old names to new names."
		list `name_old' if dup
		STOP
	}
	drop dup
	
	// store the variable names and labels as locals
	local n_vars = _N
	forvalues i = 1/`n_vars' {
		local new_`i' = `name_new'[`i']
		local old_`i' = lower(`name_old'[`i'])
		if "`label'" != "" {
			local label_`i' = `label'[`i']
		}
	}
	
	// restore the working data
	restore

	// loop over raw variables	
	foreach v of varlist _all {
	
		// label variable with the old name
		label var `v' "`v'"
	
		// rename all existing variables to all lower case 
		local lower = lower("`v'")
		cap qui rename `v' `lower'	
	}
	
	// order variables so that the ones you want to keep are first
	forvalues i = `n_vars'(-1)1 {
		order `old_`i''
	}
	rename * v_#, renumber
		
	// rename and label variables you're keeping
	forvalues i = 1/`n_vars' {
		qui cap rename v_`i' `new_`i'' 
		if "`label'" != "" {	
			label var `new_`i'' "`label_`i''"
		}
	}

	// drop extra variables if dropx option is specified
	if "`dropx'" == "dropx" {
		cap drop v_*
	}
	
	// otherwise list variables you haven't renamed yet
	else {
		local n = `n_vars' + 1
		cap confirm variable v_`n'
		if _rc == 0 {
			display as error "The following variables were not assigned a clean name.  Please provide a name for each in the supporting spreadsheet."
			describe v_*
			STOP	
		}
	}

end	
	
