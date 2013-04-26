// /Users/tamc/Documents/github/electricity-build-rate-constraint/electricity-build-rate-constraint.xlsx approximately translated into C
// First we have c versions of all the excel functions that we know
#include <stdio.h>
#include <assert.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>
#include <math.h>

// To run the tests at the end of this file
// cc excel_to_c_runtime; ./a.out

// FIXME: Extract a header file

// I predefine an array of ExcelValues to store calculations
// Probably bad practice. At the very least, I should make it
// link to the cell reference in some way.
#define MAX_EXCEL_VALUE_HEAP_SIZE 1000000
#define MAX_MEMORY_TO_BE_FREED_HEAP_SIZE 1000000

#define true 1
#define false 0

// These are the various types of excel cell, plus ExcelRange which allows the passing of arrays of cells
typedef enum {ExcelEmpty, ExcelNumber, ExcelString, ExcelBoolean, ExcelError, ExcelRange} ExcelType;

struct excel_value {
	ExcelType type;
	
	double number; // Used for numbers and for error types
	char *string; // Used for strings
	
	// The following three are used for ranges
	void *array;
	int rows;
	int columns;
};

typedef struct excel_value ExcelValue;


// These are used in the SUMIF and SUMIFS criteria (e.g., when passed a string like "<20")
typedef enum {LessThan, LessThanOrEqual, Equal, NotEqual, MoreThanOrEqual, MoreThan} ExcelComparisonType;

struct excel_comparison {
	ExcelComparisonType type;
	ExcelValue comparator;
};

typedef struct excel_comparison ExcelComparison;

// Headers
static ExcelValue more_than(ExcelValue a_v, ExcelValue b_v);
static ExcelValue more_than_or_equal(ExcelValue a_v, ExcelValue b_v);
static ExcelValue not_equal(ExcelValue a_v, ExcelValue b_v);
static ExcelValue less_than(ExcelValue a_v, ExcelValue b_v);
static ExcelValue less_than_or_equal(ExcelValue a_v, ExcelValue b_v);
static ExcelValue find_2(ExcelValue string_to_look_for_v, ExcelValue string_to_look_in_v);
static ExcelValue find(ExcelValue string_to_look_for_v, ExcelValue string_to_look_in_v, ExcelValue position_to_start_at_v);
static ExcelValue iferror(ExcelValue value, ExcelValue value_if_error);
static ExcelValue excel_index(ExcelValue array_v, ExcelValue row_number_v, ExcelValue column_number_v);
static ExcelValue excel_index_2(ExcelValue array_v, ExcelValue row_number_v);
static ExcelValue left(ExcelValue string_v, ExcelValue number_of_characters_v);
static ExcelValue left_1(ExcelValue string_v);
static ExcelValue max(int number_of_arguments, ExcelValue *arguments);
static ExcelValue min(int number_of_arguments, ExcelValue *arguments);
static ExcelValue mod(ExcelValue a_v, ExcelValue b_v);
static ExcelValue negative(ExcelValue a_v);
static ExcelValue pmt(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v);
static ExcelValue power(ExcelValue a_v, ExcelValue b_v);
static ExcelValue excel_round(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue rounddown(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue roundup(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue string_join(int number_of_arguments, ExcelValue *arguments);
static ExcelValue subtotal(ExcelValue type, int number_of_arguments, ExcelValue *arguments);
static ExcelValue sumifs(ExcelValue sum_range_v, int number_of_arguments, ExcelValue *arguments);
static ExcelValue sumif(ExcelValue check_range_v, ExcelValue criteria_v, ExcelValue sum_range_v );
static ExcelValue sumif_2(ExcelValue check_range_v, ExcelValue criteria_v);
static ExcelValue sumproduct(int number_of_arguments, ExcelValue *arguments);
static ExcelValue vlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v);
static ExcelValue vlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v, ExcelValue match_type_v);

// My little heap for excel values
ExcelValue cells[MAX_EXCEL_VALUE_HEAP_SIZE];
int cell_counter = 0;

#define HEAPCHECK if(cell_counter >= MAX_EXCEL_VALUE_HEAP_SIZE) { printf("ExcelValue heap full. Edit MAX_EXCEL_VALUE_HEAP_SIZE in the c source code."); exit(-1); }

// My little heap for keeping pointers to memory that I need to reclaim
void *memory_that_needs_to_be_freed[MAX_MEMORY_TO_BE_FREED_HEAP_SIZE];
int memory_that_needs_to_be_freed_counter = 0;

#define MEMORY_THAT_NEEDS_TO_BE_FREED_HEAP_CHECK 

static void free_later(void *pointer) {
	memory_that_needs_to_be_freed[memory_that_needs_to_be_freed_counter] = pointer;
	memory_that_needs_to_be_freed_counter++;
	if(memory_that_needs_to_be_freed_counter >= MAX_MEMORY_TO_BE_FREED_HEAP_SIZE) { 
		printf("Memory that needs to be freed heap full. Edit MAX_MEMORY_TO_BE_FREED_HEAP_SIZE in the c source code"); 
		exit(-1);
	}
}

static void free_all_allocated_memory() {
	int i;
	for(i = 0; i < memory_that_needs_to_be_freed_counter; i++) {
		free(memory_that_needs_to_be_freed[i]);
	}
	memory_that_needs_to_be_freed_counter = 0;
}

// The object initializers
static ExcelValue new_excel_number(double number) {
	cell_counter++;
	HEAPCHECK
	ExcelValue new_cell = 	cells[cell_counter];
	new_cell.type = ExcelNumber;
	new_cell.number = number;
	return new_cell;
};

static ExcelValue new_excel_string(char *string) {
	cell_counter++;
	HEAPCHECK
	ExcelValue new_cell = 	cells[cell_counter];
	new_cell.type = ExcelString;
	new_cell.string = string;
	return new_cell;
};

static ExcelValue new_excel_range(void *array, int rows, int columns) {
	cell_counter++;
	HEAPCHECK
	ExcelValue new_cell = cells[cell_counter];
	new_cell.type = ExcelRange;
	new_cell.array = array;
	new_cell.rows = rows;
	new_cell.columns = columns;
	return new_cell;
};

static void * new_excel_value_array(int size) {
	ExcelValue *pointer = malloc(sizeof(ExcelValue)*size); // Freed later
	if(pointer == 0) {
		printf("Out of memory\n");
		exit(-1);
	}
	free_later(pointer);
	return pointer;
};

// Constants
const ExcelValue BLANK = {.type = ExcelEmpty, .number = 0};

const ExcelValue ZERO = {.type = ExcelNumber, .number = 0};
const ExcelValue ONE = {.type = ExcelNumber, .number = 1};
const ExcelValue TWO = {.type = ExcelNumber, .number = 2};
const ExcelValue THREE = {.type = ExcelNumber, .number = 3};
const ExcelValue FOUR = {.type = ExcelNumber, .number = 4};
const ExcelValue FIVE = {.type = ExcelNumber, .number = 5};
const ExcelValue SIX = {.type = ExcelNumber, .number = 6};
const ExcelValue SEVEN = {.type = ExcelNumber, .number = 7};
const ExcelValue EIGHT = {.type = ExcelNumber, .number = 8};
const ExcelValue NINE = {.type = ExcelNumber, .number = 9};
const ExcelValue TEN = {.type = ExcelNumber, .number = 10};

// Booleans
const ExcelValue TRUE = {.type = ExcelBoolean, .number = true };
const ExcelValue FALSE = {.type = ExcelBoolean, .number = false };

// Errors
const ExcelValue VALUE = {.type = ExcelError, .number = 0};
const ExcelValue NAME = {.type = ExcelError, .number = 1};
const ExcelValue DIV0 = {.type = ExcelError, .number = 2};
const ExcelValue REF = {.type = ExcelError, .number = 3};
const ExcelValue NA = {.type = ExcelError, .number = 4};

// This is the error flag
static int conversion_error = 0;

// Helpful for debugging
static void inspect_excel_value(ExcelValue v) {
	ExcelValue *array;
	int i, j, k;
	switch (v.type) {
  	  case ExcelNumber:
		  printf("Number: %f\n",v.number);
		  break;
	  case ExcelBoolean:
		  if(v.number == true) {
			  printf("True\n");
		  } else if(v.number == false) {
			  printf("False\n");
		  } else {
			  printf("Boolean with undefined state %f\n",v.number);
		  }
		  break;
	  case ExcelEmpty:
	  	if(v.number == 0) {
	  		printf("Empty\n");
		} else {
			printf("Empty with unexpected state %f\n",v.number);	
		}
		break;
	  case ExcelRange:
		 printf("Range rows: %d, columns: %d\n",v.rows,v.columns);
		 array = v.array;
		 for(i = 0; i < v.rows; i++) {
			 printf("Row %d:\n",i+1);
			 for(j = 0; j < v.columns; j++ ) {
				 printf("%d ",j+1);
				 k = (i * v.columns) + j;
				 inspect_excel_value(array[k]);
			 }
		 }
		 break;
	  case ExcelString:
		 printf("String: '%s'\n",v.string);
		 break;
	  case ExcelError:
		 printf("Error number %f ",v.number);
		 switch( (int)v.number) {
			 case 0: printf("VALUE\n"); break;
			 case 1: printf("NAME\n"); break;
			 case 2: printf("DIV0\n"); break;
			 case 3: printf("REF\n"); break;
			 case 4: printf("NA\n"); break;
		 }
		 break;
    default:
      printf("Type %d not recognised",v.type);
	 };
}

// Extracts numbers from ExcelValues
// Excel treats empty cells as zero
static double number_from(ExcelValue v) {
	char *s;
	char *p;
	double n;
	ExcelValue *array;
	switch (v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  	return v.number;
	  case ExcelEmpty: 
	  	return 0;
	  case ExcelRange: 
		 array = v.array;
	     return number_from(array[0]);
	  case ExcelString:
 	 	s = v.string;
		if (s == NULL || *s == '\0' || isspace(*s)) {
			return 0;
		}	        
		n = strtod (s, &p);
		if(*p == '\0') {
			return n;
		}
		conversion_error = 1;
		return 0;
	  case ExcelError:
	  	return 0;
  }
  return 0;
}

#define NUMBER(value_name, name) double name; if(value_name.type == ExcelError) { return value_name; }; name = number_from(value_name);
#define CHECK_FOR_CONVERSION_ERROR 	if(conversion_error) { conversion_error = 0; return VALUE; };
#define CHECK_FOR_PASSED_ERROR(name) 	if(name.type == ExcelError) return name;
	
static ExcelValue excel_abs(ExcelValue a_v) {
	CHECK_FOR_PASSED_ERROR(a_v)	
	NUMBER(a_v, a)
	CHECK_FOR_CONVERSION_ERROR
	
	if(a >= 0.0 ) {
		return a_v;
	} else {
		return new_excel_number(-a);
	}
}

static ExcelValue add(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a + b);
}

static ExcelValue excel_and(int array_size, ExcelValue *array) {
	int i;
	ExcelValue current_excel_value, array_result;
	
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		switch (current_excel_value.type) {
	  	  case ExcelNumber: 
		  case ExcelBoolean: 
			  if(current_excel_value.number == false) return FALSE;
			  break;
		  case ExcelRange: 
		  	array_result = excel_and( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(array_result.type == ExcelError) return array_result;
			if(array_result.type == ExcelBoolean && array_result.number == false) return FALSE;
			break;
		  case ExcelString:
		  case ExcelEmpty:
			 break;
		  case ExcelError:
			 return current_excel_value;
			 break;
		 }
	 }
	 return TRUE;
}

struct average_result {
	double sum;
	double count;
	int has_error;
	ExcelValue error;
};
	
static struct average_result calculate_average(int array_size, ExcelValue *array) {
	double sum = 0;
	double count = 0;
	int i;
	ExcelValue current_excel_value;
	struct average_result array_result, r;
		 
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		switch (current_excel_value.type) {
	  	  case ExcelNumber:
			  sum += current_excel_value.number;
			  count++;
			  break;
		  case ExcelRange: 
		  	array_result = calculate_average( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(array_result.has_error == true) return array_result;
			sum += array_result.sum;
			count += array_result.count;
			break;
		  case ExcelBoolean: 
		  case ExcelString:
		  case ExcelEmpty:
			 break;
		  case ExcelError:
			 r.has_error = true;
			 r.error = current_excel_value;
			 return r;
			 break;
		 }
	}
	r.count = count;
	r.sum = sum;
	r.has_error = false;
	return r;
}

static ExcelValue average(int array_size, ExcelValue *array) {
	struct average_result r = calculate_average(array_size, array);
	if(r.has_error == true) return r.error;
	if(r.count == 0) return DIV0;
	return new_excel_number(r.sum/r.count);
}

static ExcelValue choose(ExcelValue index_v, int array_size, ExcelValue *array) {
	CHECK_FOR_PASSED_ERROR(index_v)

	int index = (int) number_from(index_v);
	CHECK_FOR_CONVERSION_ERROR	
	int i;
	for(i=0;i<array_size;i++) {
		if(array[i].type == ExcelError) return array[i];
	}
	if(index < 1) return VALUE;
	if(index > array_size) return VALUE;
	return array[index-1];
}	

static ExcelValue count(int array_size, ExcelValue *array) {
	int i;
	int n = 0;
	ExcelValue current_excel_value;
	
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		switch (current_excel_value.type) {
	  	  case ExcelNumber:
		  	n++;
			break;
		  case ExcelRange: 
		  	n += count( current_excel_value.rows * current_excel_value.columns, current_excel_value.array ).number;
			break;
  		  case ExcelBoolean: 			
		  case ExcelString:
		  case ExcelEmpty:
		  case ExcelError:
			 break;
		 }
	 }
	 return new_excel_number(n);
}

static ExcelValue counta(int array_size, ExcelValue *array) {
	int i;
	int n = 0;
	ExcelValue current_excel_value;
	
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
    switch(current_excel_value.type) {
  	  case ExcelNumber:
      case ExcelBoolean:
      case ExcelString:
  	  case ExcelError:
        n++;
        break;
      case ExcelRange: 
	  	  n += counta( current_excel_value.rows * current_excel_value.columns, current_excel_value.array ).number;
        break;
  	  case ExcelEmpty:
  		  break;
    }
	 }
	 return new_excel_number(n);
}

static ExcelValue divide(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	if(b == 0) return DIV0;
	return new_excel_number(a / b);
}

static ExcelValue excel_equal(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	if(a_v.type != b_v.type) return FALSE;
	
	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty: 
			if(a_v.number != b_v.number) return FALSE;
			return TRUE;
	  case ExcelString:
	  	if(strcasecmp(a_v.string,b_v.string) != 0 ) return FALSE;
		return TRUE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}

static ExcelValue not_equal(ExcelValue a_v, ExcelValue b_v) {
	ExcelValue result = excel_equal(a_v, b_v);
	if(result.type == ExcelBoolean) {
		if(result.number == 0) return TRUE;
		return FALSE;
	}
	return result;
}

static ExcelValue excel_if(ExcelValue condition, ExcelValue true_case, ExcelValue false_case ) {
	CHECK_FOR_PASSED_ERROR(condition)
	
	switch (condition.type) {
  	  case ExcelBoolean:
  	  	if(condition.number == true) return true_case;
  	  	return false_case;
  	  case ExcelNumber:
		if(condition.number == false) return false_case;
		return true_case;
	  case ExcelEmpty: 
		return false_case;
	  case ExcelString:
	  	return VALUE;
  	  case ExcelError:
		return condition;
  	  case ExcelRange:
  		return VALUE;
  }
  return condition;
}

static ExcelValue excel_if_2(ExcelValue condition, ExcelValue true_case ) {
	return excel_if( condition, true_case, FALSE );
}

static ExcelValue excel_index(ExcelValue array_v, ExcelValue row_number_v, ExcelValue column_number_v) {
	CHECK_FOR_PASSED_ERROR(array_v)
	CHECK_FOR_PASSED_ERROR(row_number_v)
	CHECK_FOR_PASSED_ERROR(column_number_v)
		
	ExcelValue *array;
	int rows;
	int columns;
	
	NUMBER(row_number_v, row_number)
	NUMBER(column_number_v, column_number)
	CHECK_FOR_CONVERSION_ERROR
	
	if(array_v.type == ExcelRange) {
		array = array_v.array;
		rows = array_v.rows;
		columns = array_v.columns;
	} else {
		ExcelValue tmp_array[] = {array_v};
		array = tmp_array;
		rows = 1;
		columns = 1;
	}
	
	if(row_number > rows) return REF;
	if(column_number > columns) return REF;
		
	if(row_number == 0) { // We need the whole column
		if(column_number < 1) return REF;
		ExcelValue *result = (ExcelValue *) new_excel_value_array(rows);
		int result_index = 0;
		ExcelValue r;
		int array_index;
		int i;
		for(i = 0; i < rows; i++) {
			array_index = (i*columns) + column_number - 1;
			r = array[array_index];
			if(r.type == ExcelEmpty) {
				result[result_index] = ZERO;
			} else {
				result[result_index] = r;
			}			
			result_index++;
		}
		return new_excel_range(result,rows,1);
	} else if(column_number == 0 ) { // We need the whole row
		if(row_number < 1) return REF;
		ExcelValue *result = (ExcelValue*) new_excel_value_array(columns);
		ExcelValue r;
		int row_start = ((row_number-1)*columns);
		int row_finish = row_start + columns;
		int result_index = 0;
		int i;
		for(i = row_start; i < row_finish; i++) {
			r = array[i];
			if(r.type == ExcelEmpty) {
				result[result_index] = ZERO;
			} else {
				result[result_index] = r;
			}
			result_index++;
		}
		return new_excel_range(result,1,columns);
	} else { // We need a precise point
		if(row_number < 1) return REF;
		if(column_number < 1) return REF;
		int position = ((row_number - 1) * columns) + column_number - 1;
		ExcelValue result = array[position];
		if(result.type == ExcelEmpty) return ZERO;
		return result;
	}
	
	return FALSE;
};

static ExcelValue excel_index_2(ExcelValue array_v, ExcelValue offset) {
	if(array_v.type == ExcelRange) {
		if(array_v.rows == 1) {
			return excel_index(array_v,ONE,offset);
		} else if (array_v.columns == 1) {
			return excel_index(array_v,offset,ONE);
		} else {
			return REF;
		}
	} else if (offset.type == ExcelNumber && offset.number == 1) {
		return array_v;
	} else {
		return REF;
	}
	return REF;
};


static ExcelValue excel_match(ExcelValue lookup_value, ExcelValue lookup_array, ExcelValue match_type ) {
	CHECK_FOR_PASSED_ERROR(lookup_value)
	CHECK_FOR_PASSED_ERROR(lookup_array)
	CHECK_FOR_PASSED_ERROR(match_type)
		
	// Blanks are treaked as zeros
	if(lookup_value.type == ExcelEmpty) lookup_value = ZERO;

	// Setup the array
	ExcelValue *array;
	int size;
	if(lookup_array.type == ExcelRange) {
		// Check that the range is a row or column rather than an area
		if((lookup_array.rows == 1) || (lookup_array.columns == 1)) {
			array = lookup_array.array;
			size = lookup_array.rows * lookup_array.columns;
		} else {
			// return NA error if covers an area.
			return NA;
		};
	} else {
		// Need to wrap the argument up as an array
		size = 1;
		ExcelValue tmp_array[1] = {lookup_array};
		array = tmp_array;
	}
    
	int type = (int) number_from(match_type);
	CHECK_FOR_CONVERSION_ERROR;
	
	int i;
	ExcelValue x;
	
	switch(type) {
		case 0:
			for(i = 0; i < size; i++ ) {
				x = array[i];
				if(x.type == ExcelEmpty) x = ZERO;
				if(excel_equal(lookup_value,x).number == true) return new_excel_number(i+1);
			}
			return NA;
			break;
		case 1:
			for(i = 0; i < size; i++ ) {
				x = array[i];
				if(x.type == ExcelEmpty) x = ZERO;
				if(more_than(x,lookup_value).number == true) {
					if(i==0) return NA;
					return new_excel_number(i);
				}
			}
			return new_excel_number(size);
			break;
		case -1:
			for(i = 0; i < size; i++ ) {
				x = array[i];
				if(x.type == ExcelEmpty) x = ZERO;
				if(less_than(x,lookup_value).number == true) {
					if(i==0) return NA;
					return new_excel_number(i);
				}
			}
			return new_excel_number(size-1);
			break;
	}
	return NA;
}

static ExcelValue excel_match_2(ExcelValue lookup_value, ExcelValue lookup_array ) {
	return excel_match(lookup_value,lookup_array,new_excel_number(0));
}

static ExcelValue find(ExcelValue find_text_v, ExcelValue within_text_v, ExcelValue start_number_v) {
	CHECK_FOR_PASSED_ERROR(find_text_v)
	CHECK_FOR_PASSED_ERROR(within_text_v)
	CHECK_FOR_PASSED_ERROR(start_number_v)

	char *find_text;	
	char *within_text;
	char *within_text_offset;
	char *result;
	int start_number = number_from(start_number_v);
	CHECK_FOR_CONVERSION_ERROR

	// Deal with blanks 
	if(within_text_v.type == ExcelString) {
		within_text = within_text_v.string;
	} else if( within_text_v.type == ExcelEmpty) {
		within_text = "";
	}

	if(find_text_v.type == ExcelString) {
		find_text = find_text_v.string;
	} else if( find_text_v.type == ExcelEmpty) {
		return start_number_v;
	}
	
	// Check length
	if(start_number < 1) return VALUE;
	if(start_number > strlen(within_text)) return VALUE;
	
	// Offset our within_text pointer
	// FIXME: No way this is utf-8 compatible
	within_text_offset = within_text + (start_number - 1);
	result = strstr(within_text_offset,find_text);
	if(result) {
		return new_excel_number(result - within_text + 1);
	}
	return VALUE;
}

static ExcelValue find_2(ExcelValue string_to_look_for_v, ExcelValue string_to_look_in_v) {
	return find(string_to_look_for_v, string_to_look_in_v, ONE);
};

static ExcelValue left(ExcelValue string_v, ExcelValue number_of_characters_v) {
	CHECK_FOR_PASSED_ERROR(string_v)
	CHECK_FOR_PASSED_ERROR(number_of_characters_v)
	if(string_v.type == ExcelEmpty) return BLANK;
	if(number_of_characters_v.type == ExcelEmpty) return BLANK;
	
	int number_of_characters = (int) number_from(number_of_characters_v);
	CHECK_FOR_CONVERSION_ERROR

	char *string;
	int string_must_be_freed = 0;
	switch (string_v.type) {
  	  case ExcelString:
  		string = string_v.string;
  		break;
  	  case ExcelNumber:
		  string = malloc(20); // Freed
		  if(string == 0) {
			  printf("Out of memory");
			  exit(-1);
		  }
		  string_must_be_freed = 1;
		  snprintf(string,20,"%f",string_v.number);
		  break;
	  case ExcelBoolean:
	  	if(string_v.number == true) {
	  		string = "TRUE";
		} else {
			string = "FALSE";
		}
		break;
	  case ExcelEmpty:	  	 
  	  case ExcelError:
  	  case ExcelRange:
		return string_v;
	}
	
	char *left_string = malloc(number_of_characters+1); // Freed
	if(left_string == 0) {
	  printf("Out of memory");
	  exit(-1);
	}
	free_later(left_string);
	memcpy(left_string,string,number_of_characters);
	left_string[number_of_characters] = '\0';
	if(string_must_be_freed == 1) {
		free(string);
	}
	return new_excel_string(left_string);
}

static ExcelValue left_1(ExcelValue string_v) {
	return left(string_v, ONE);
}

static ExcelValue iferror(ExcelValue value, ExcelValue value_if_error) {
	if(value.type == ExcelError) return value_if_error;
	return value;
}

static ExcelValue more_than(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty:
		if((b_v.type == ExcelNumber) || (b_v.type == ExcelBoolean) || (b_v.type == ExcelEmpty)) {
			if(a_v.number <= b_v.number) return FALSE;
			return TRUE;
		} 
		return FALSE;
	  case ExcelString:
	  	if(b_v.type == ExcelString) {
		  	if(strcasecmp(a_v.string,b_v.string) <= 0 ) return FALSE;
			return TRUE;	  		
		}
		return FALSE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}

static ExcelValue more_than_or_equal(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty:
		if((b_v.type == ExcelNumber) || (b_v.type == ExcelBoolean) || (b_v.type == ExcelEmpty)) {
			if(a_v.number < b_v.number) return FALSE;
			return TRUE;
		} 
		return FALSE;
	  case ExcelString:
	  	if(b_v.type == ExcelString) {
		  	if(strcasecmp(a_v.string,b_v.string) < 0 ) return FALSE;
			return TRUE;	  		
		}
		return FALSE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}


static ExcelValue less_than(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty:
		if((b_v.type == ExcelNumber) || (b_v.type == ExcelBoolean) || (b_v.type == ExcelEmpty)) {
			if(a_v.number >= b_v.number) return FALSE;
			return TRUE;
		} 
		return FALSE;
	  case ExcelString:
	  	if(b_v.type == ExcelString) {
		  	if(strcasecmp(a_v.string,b_v.string) >= 0 ) return FALSE;
			return TRUE;	  		
		}
		return FALSE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}

static ExcelValue less_than_or_equal(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty:
		if((b_v.type == ExcelNumber) || (b_v.type == ExcelBoolean) || (b_v.type == ExcelEmpty)) {
			if(a_v.number > b_v.number) return FALSE;
			return TRUE;
		} 
		return FALSE;
	  case ExcelString:
	  	if(b_v.type == ExcelString) {
		  	if(strcasecmp(a_v.string,b_v.string) > 0 ) return FALSE;
			return TRUE;	  		
		}
		return FALSE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}

static ExcelValue subtract(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a - b);
}

static ExcelValue multiply(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a * b);
}

static ExcelValue sum(int array_size, ExcelValue *array) {
	double total = 0;
	int i;
	for(i=0;i<array_size;i++) {
    switch(array[i].type) {
      case ExcelNumber:
        total += array[i].number;
        break;
      case ExcelRange:
        total += number_from(sum( array[i].rows * array[i].columns, array[i].array ));
        break;
      case ExcelError:
        return array[i];
        break;
      default:
        break;
    }
	}
	return new_excel_number(total);
}

static ExcelValue max(int number_of_arguments, ExcelValue *arguments) {
	double biggest_number_found;
	int any_number_found = 0;
	int i;
	ExcelValue current_excel_value;
	
	for(i=0;i<number_of_arguments;i++) {
		current_excel_value = arguments[i];
		if(current_excel_value.type == ExcelNumber) {
			if(!any_number_found) {
				any_number_found = 1;
				biggest_number_found = current_excel_value.number;
			}
			if(current_excel_value.number > biggest_number_found) biggest_number_found = current_excel_value.number; 				
		} else if(current_excel_value.type == ExcelRange) {
			current_excel_value = max( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(current_excel_value.type == ExcelError) return current_excel_value;
			if(current_excel_value.type == ExcelNumber)
				if(!any_number_found) {
					any_number_found = 1;
					biggest_number_found = current_excel_value.number;
				}
				if(current_excel_value.number > biggest_number_found) biggest_number_found = current_excel_value.number; 				
		} else if(current_excel_value.type == ExcelError) {
			return current_excel_value;
		}
	}
	if(!any_number_found) {
		any_number_found = 1;
		biggest_number_found = 0;
	}
	return new_excel_number(biggest_number_found);	
}

static ExcelValue min(int number_of_arguments, ExcelValue *arguments) {
	double smallest_number_found = 0;
	int any_number_found = 0;
	int i;
	ExcelValue current_excel_value;
	
	for(i=0;i<number_of_arguments;i++) {
		current_excel_value = arguments[i];
		if(current_excel_value.type == ExcelNumber) {
			if(!any_number_found) {
				any_number_found = 1;
				smallest_number_found = current_excel_value.number;
			}
			if(current_excel_value.number < smallest_number_found) smallest_number_found = current_excel_value.number; 				
		} else if(current_excel_value.type == ExcelRange) {
			current_excel_value = min( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(current_excel_value.type == ExcelError) return current_excel_value;
			if(current_excel_value.type == ExcelNumber)
				if(!any_number_found) {
					any_number_found = 1;
					smallest_number_found = current_excel_value.number;
				}
				if(current_excel_value.number < smallest_number_found) smallest_number_found = current_excel_value.number; 				
		} else if(current_excel_value.type == ExcelError) {
			return current_excel_value;
		}
	}
	if(!any_number_found) {
		any_number_found = 1;
		smallest_number_found = 0;
	}
	return new_excel_number(smallest_number_found);	
}

static ExcelValue mod(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
		
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	if(b == 0) return DIV0;
	return new_excel_number(fmod(a,b));
}

static ExcelValue negative(ExcelValue a_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	NUMBER(a_v, a)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(-a);
}

static ExcelValue pmt(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v) {
	CHECK_FOR_PASSED_ERROR(rate_v)
	CHECK_FOR_PASSED_ERROR(number_of_periods_v)
	CHECK_FOR_PASSED_ERROR(present_value_v)
		
	NUMBER(rate_v,rate)
	NUMBER(number_of_periods_v,number_of_periods)
	NUMBER(present_value_v,present_value)
	CHECK_FOR_CONVERSION_ERROR
	
	if(rate == 0) return new_excel_number(-(present_value / number_of_periods));
	return new_excel_number(-present_value*(rate*(pow((1+rate),number_of_periods)))/((pow((1+rate),number_of_periods))-1));
}

static ExcelValue power(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
		
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(pow(a,b));
}

static ExcelValue excel_round(ExcelValue number_v, ExcelValue decimal_places_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(decimal_places_v)
		
	NUMBER(number_v, number)
	NUMBER(decimal_places_v, decimal_places)
	CHECK_FOR_CONVERSION_ERROR
		
	double multiple = pow(10,decimal_places);
	
	return new_excel_number( round(number * multiple) / multiple );
}

static ExcelValue rounddown(ExcelValue number_v, ExcelValue decimal_places_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(decimal_places_v)
		
	NUMBER(number_v, number)
	NUMBER(decimal_places_v, decimal_places)
	CHECK_FOR_CONVERSION_ERROR
		
	double multiple = pow(10,decimal_places);
	
	return new_excel_number( trunc(number * multiple) / multiple );	
}

static ExcelValue roundup(ExcelValue number_v, ExcelValue decimal_places_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(decimal_places_v)
		
	NUMBER(number_v, number)
	NUMBER(decimal_places_v, decimal_places)
	CHECK_FOR_CONVERSION_ERROR
		
	double multiple = pow(10,decimal_places);
	if(number < 0) return new_excel_number( floor(number * multiple) / multiple );
	return new_excel_number( ceil(number * multiple) / multiple );	
}

static ExcelValue string_join(int number_of_arguments, ExcelValue *arguments) {
	int allocated_length = 100;
	int used_length = 0;
	char *string = malloc(allocated_length); // Freed later
	if(string == 0) {
	  printf("Out of memory");
	  exit(-1);
	}
	free_later(string);
	char *current_string;
	int current_string_length;
	int must_free_current_string;
	ExcelValue current_v;
	int i;
	for(i=0;i<number_of_arguments;i++) {
		must_free_current_string = 0;
		current_v = (ExcelValue) arguments[i];
		switch (current_v.type) {
  	  case ExcelString:
	  		current_string = current_v.string;
	  		break;
  	  case ExcelNumber:
			  current_string = malloc(20); // Freed
		  	if(current_string == 0) {
		  	  printf("Out of memory");
		  	  exit(-1);
		  	}
			must_free_current_string = 1;				  
			  snprintf(current_string,20,"%g",current_v.number);
			  break;
		  case ExcelBoolean:
		  	if(current_v.number == true) {
		  		current_string = "TRUE";
  			} else {
  				current_string = "FALSE";
  			}
        break;
		  case ExcelEmpty:
        current_string = "";
        break;
      case ExcelError:
        return current_v;
	  	case ExcelRange:
        return VALUE;
		}
		current_string_length = strlen(current_string);
		if( (used_length + current_string_length + 1) > allocated_length) {
			allocated_length += 100;
			string = realloc(string,allocated_length);
		}
		memcpy(string + used_length, current_string, current_string_length);
		if(must_free_current_string == 1) {
			free(current_string);
		}
		used_length = used_length + current_string_length;
	}
	string = realloc(string,used_length+1);
  string[used_length] = '\0';
	return new_excel_string(string);
}

static ExcelValue subtotal(ExcelValue subtotal_type_v, int number_of_arguments, ExcelValue *arguments) {
  CHECK_FOR_PASSED_ERROR(subtotal_type_v)
  NUMBER(subtotal_type_v,subtotal_type)
  CHECK_FOR_CONVERSION_ERROR
      
  switch((int) subtotal_type) {
    case 1:
    case 101:
      return average(number_of_arguments,arguments);
      break;
    case 2:
    case 102:
      return count(number_of_arguments,arguments);
      break;
    case 3:
    case 103:
      return counta(number_of_arguments,arguments);
      break;
    case 9:
    case 109:
      return sum(number_of_arguments,arguments);
      break;
    default:
      return VALUE;
      break;
  }
}

static ExcelValue sumifs(ExcelValue sum_range_v, int number_of_arguments, ExcelValue *arguments) {
  // First, set up the sum_range
  CHECK_FOR_PASSED_ERROR(sum_range_v);

  // Set up the sum range
  ExcelValue *sum_range;
  int sum_range_rows, sum_range_columns;
  
  if(sum_range_v.type == ExcelRange) {
    sum_range = sum_range_v.array;
    sum_range_rows = sum_range_v.rows;
    sum_range_columns = sum_range_v.columns;
  } else {
    sum_range = (ExcelValue*) new_excel_value_array(1);
	sum_range[0] = sum_range_v;
    sum_range_rows = 1;
    sum_range_columns = 1;
  }
  
  // Then go through and set up the check ranges
  if(number_of_arguments % 2 != 0) return VALUE;
  int number_of_criteria = number_of_arguments / 2;
  ExcelValue *criteria_range =  (ExcelValue*) new_excel_value_array(number_of_criteria);
  ExcelValue current_value;
  int i;
  for(i = 0; i < number_of_criteria; i++) {
    current_value = arguments[i*2];
    if(current_value.type == ExcelRange) {
      criteria_range[i] = current_value;
      if(current_value.rows != sum_range_rows) return VALUE;
      if(current_value.columns != sum_range_columns) return VALUE;
    } else {
      if(sum_range_rows != 1) return VALUE;
      if(sum_range_columns != 1) return VALUE;
      ExcelValue *tmp_array2 =  (ExcelValue*) new_excel_value_array(1);
      tmp_array2[0] = current_value;
      criteria_range[i] =  new_excel_range(tmp_array2,1,1);
    }
  }
  
  // Now go through and set up the criteria
  ExcelComparison *criteria =  malloc(sizeof(ExcelComparison)*number_of_criteria); // freed at end of function
  if(criteria == 0) {
	  printf("Out of memory\n");
	  exit(-1);
  }
  char *s;
  for(i = 0; i < number_of_criteria; i++) {
    current_value = arguments[(i*2)+1];
    
    if(current_value.type == ExcelString) {
      s = current_value.string;
      if(s[0] == '<') {
        if( s[1] == '>') {
          criteria[i].type = NotEqual;
          criteria[i].comparator = new_excel_string(strndup(s+2,strlen(s)-2));
        } else if(s[1] == '=') {
          criteria[i].type = LessThanOrEqual;
          criteria[i].comparator = new_excel_string(strndup(s+2,strlen(s)-2));
        } else {
          criteria[i].type = LessThan;
          criteria[i].comparator = new_excel_string(strndup(s+1,strlen(s)-1));
        }
      } else if(s[0] == '>') {
        if(s[1] == '=') {
          criteria[i].type = MoreThanOrEqual;
          criteria[i].comparator = new_excel_string(strndup(s+2,strlen(s)-2));
        } else {
          criteria[i].type = MoreThan;
          criteria[i].comparator = new_excel_string(strndup(s+1,strlen(s)-1));
        }
      } else if(s[0] == '=') {
        criteria[i].type = Equal;
        criteria[i].comparator = new_excel_string(strndup(s+1,strlen(s)-1));          
      } else {
        criteria[i].type = Equal;
        criteria[i].comparator = current_value;          
      }
    } else {
      criteria[i].type = Equal;
      criteria[i].comparator = current_value;
    }
  }
  
  double total = 0;
  int size = sum_range_columns * sum_range_rows;
  int j;
  int passed = 0;
  ExcelValue value_to_be_checked;
  ExcelComparison comparison;
  ExcelValue comparator;
  double number;
  // For each cell in the sum range
  for(j=0; j < size; j++ ) {
    passed = 1;
    for(i=0; i < number_of_criteria; i++) {
      value_to_be_checked = ((ExcelValue *) ((ExcelValue) criteria_range[i]).array)[j];
      comparison = criteria[i];
      comparator = comparison.comparator;
      
      switch(value_to_be_checked.type) {
        case ExcelError: // Errors match only errors
          if(comparison.type != Equal) passed = 0;
          if(comparator.type != ExcelError) passed = 0;
          if(value_to_be_checked.number != comparator.number) passed = 0;
          break;
        case ExcelBoolean: // Booleans match only booleans (FIXME: I think?)
          if(comparison.type != Equal) passed = 0;
          if(comparator.type != ExcelBoolean ) passed = 0;
          if(value_to_be_checked.number != comparator.number) passed = 0;
          break;
        case ExcelEmpty:
          // if(comparator.type == ExcelEmpty) break; // FIXME: Huh? In excel blank doesn't match blank?!
          if(comparator.type != ExcelString) {
            passed = 0;
            break;
          } else {
            if(strlen(comparator.string) != 0) passed = 0; // Empty strings match blanks.
            break;
          }
          break;
        case ExcelNumber:
          if(comparator.type == ExcelNumber) {
            number = comparator.number;
          } else if(comparator.type == ExcelString) {
            number = number_from(comparator);
            if(conversion_error == 1) {
              conversion_error = 0;
              passed = 0;
              break;
            }
          } else {
            passed = 0;
            break;
          }
          switch(comparison.type) {
            case Equal:
              if(value_to_be_checked.number != number) passed = 0;
              break;
            case LessThan:
              if(value_to_be_checked.number >= number) passed = 0;
              break;            
            case LessThanOrEqual:
              if(value_to_be_checked.number > number) passed = 0;
              break;                        
            case NotEqual:
              if(value_to_be_checked.number == number) passed = 0;
              break;            
            case MoreThanOrEqual:
              if(value_to_be_checked.number < number) passed = 0;
              break;            
            case MoreThan:
              if(value_to_be_checked.number <= number) passed = 0;
              break;
          }
          break;
        case ExcelString:
          // First case, the comparator is a number, simplification is that it can only be equal
          if(comparator.type == ExcelNumber) {
            if(comparison.type != Equal) {
              printf("This shouldn't be possible?");
              passed = 0;
              break;
            }
            number = number_from(value_to_be_checked);
            if(conversion_error == 1) {
              conversion_error = 0;
              passed = 0;
              break;
            }
            if(number != comparator.number) {
              passed = 0;
              break;
            } else {
              break;
            }
          // Second case, the comparator is also a string, so need to be able to do full range of tests
          } else if(comparator.type == ExcelString) {
            switch(comparison.type) {
              case Equal:
                if(excel_equal(value_to_be_checked,comparator).number == 0) passed = 0;
                break;
              case LessThan:
                if(less_than(value_to_be_checked,comparator).number == 0) passed = 0;
                break;            
              case LessThanOrEqual:
                if(less_than_or_equal(value_to_be_checked,comparator).number == 0) passed = 0;
                break;                        
              case NotEqual:
                if(not_equal(value_to_be_checked,comparator).number == 0) passed = 0;
                break;            
              case MoreThanOrEqual:
                if(more_than_or_equal(value_to_be_checked,comparator).number == 0) passed = 0;
                break;            
              case MoreThan:
                if(more_than(value_to_be_checked,comparator).number == 0) passed = 0;
                break;
              }
          } else {
            passed = 0;
            break;
          }
          break;
        case ExcelRange:
          return VALUE;            
      }
      if(passed == 0) break;
    }
    if(passed == 1) {
      current_value = sum_range[j];
      if(current_value.type == ExcelError) {
        return current_value;
      } else if(current_value.type == ExcelNumber) {
        total += current_value.number;
      }
    }
  }
  // Tidy up
  free(criteria);
  return new_excel_number(total);
}

static ExcelValue sumif(ExcelValue check_range_v, ExcelValue criteria_v, ExcelValue sum_range_v ) {
	ExcelValue tmp_array_sumif[] = {check_range_v, criteria_v};
	return sumifs(sum_range_v,2,tmp_array_sumif);
}

static ExcelValue sumif_2(ExcelValue check_range_v, ExcelValue criteria_v) {
	ExcelValue tmp_array_sumif2[] = {check_range_v, criteria_v};
	return sumifs(check_range_v,2,tmp_array_sumif2);
}

static ExcelValue sumproduct(int number_of_arguments, ExcelValue *arguments) {
  if(number_of_arguments <1) return VALUE;
  
  int a;
  int i;
  int j;
  int rows;
  int columns;
  ExcelValue current_value;
  ExcelValue **ranges = malloc(sizeof(ExcelValue *)*number_of_arguments); // Added free statements
  if(ranges == 0) {
	  printf("Out of memory\n");
	  exit(-1);
  }
  double product = 1;
  double sum = 0;
  
  // Find out dimensions of first argument
  if(arguments[0].type == ExcelRange) {
    rows = arguments[0].rows;
    columns = arguments[0].columns;
  } else {
    rows = 1;
    columns = 1;
  }
  // Extract arrays from each of the given ranges, checking for errors and bounds as we go
  for(a=0;a<number_of_arguments;a++) {
    current_value = arguments[a];
    switch(current_value.type) {
      case ExcelRange:
        if(current_value.rows != rows || current_value.columns != columns) return VALUE;
        ranges[a] = current_value.array;
        break;
      case ExcelError:
		free(ranges);
        return current_value;
        break;
      case ExcelEmpty:
		free(ranges);
        return VALUE;
        break;
      default:
        if(rows != 1 && columns !=1) return VALUE;
        ranges[a] = (ExcelValue*) new_excel_value_array(1);
        ranges[a][0] = arguments[a];
        break;
    }
  }
  	
	for(i=0;i<rows;i++) {
		for(j=0;j<columns;j++) {
			product = 1;
			for(a=0;a<number_of_arguments;a++) {
				current_value = ranges[a][(i*columns)+j];
				if(current_value.type == ExcelNumber) {
					product *= current_value.number;
				} else {
					product *= 0;
				}
			}
			sum += product;
		}
	}
	free(ranges);
  	return new_excel_number(sum);
}

static ExcelValue vlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v) {
  return vlookup(lookup_value_v,lookup_table_v,column_number_v,TRUE);
}

static ExcelValue vlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v, ExcelValue match_type_v) {
  CHECK_FOR_PASSED_ERROR(lookup_value_v)
  CHECK_FOR_PASSED_ERROR(lookup_table_v)
  CHECK_FOR_PASSED_ERROR(column_number_v)
  CHECK_FOR_PASSED_ERROR(match_type_v)

  if(lookup_value_v.type == ExcelEmpty) return NA;
  if(lookup_table_v.type != ExcelRange) return NA;
  if(column_number_v.type != ExcelNumber) return NA;
  if(match_type_v.type != ExcelBoolean) return NA;
    
  int i;
  int last_good_match = 0;
  int rows = lookup_table_v.rows;
  int columns = lookup_table_v.columns;
  ExcelValue *array = lookup_table_v.array;
  ExcelValue possible_match_v;
  
  if(column_number_v.number > columns) return REF;
  if(column_number_v.number < 1) return VALUE;
  
  if(match_type_v.number == false) { // Exact match required
    for(i=0; i< rows; i++) {
      possible_match_v = array[i*columns];
      if(excel_equal(lookup_value_v,possible_match_v).number == true) {
        return array[(i*columns)+(((int) column_number_v.number) - 1)];
      }
    }
    return NA;
  } else { // Highest value that is less than or equal
    for(i=0; i< rows; i++) {
      possible_match_v = array[i*columns];
      if(lookup_value_v.type != possible_match_v.type) continue;
      if(more_than(possible_match_v,lookup_value_v).number == true) {
        if(i == 0) return NA;
        return array[((i-1)*columns)+(((int) column_number_v.number) - 1)];
      } else {
        last_good_match = i;
      }
    }
    return array[(last_good_match*columns)+(((int) column_number_v.number) - 1)];   
  }
  return NA;
}



int test_functions() {
	// Test ABS
	assert(excel_abs(ONE).number == 1);
	assert(excel_abs(new_excel_number(-1)).number == 1);
	assert(excel_abs(VALUE).type == ExcelError);
	
	// Test ADD
	assert(add(ONE,new_excel_number(-2.5)).number == -1.5);
	assert(add(ONE,VALUE).type == ExcelError);
	
	// Test AND
	ExcelValue true_array1[] = { TRUE, new_excel_number(10)};
	ExcelValue true_array2[] = { ONE };
	ExcelValue false_array1[] = { FALSE, new_excel_number(10)};
	ExcelValue false_array2[] = { TRUE, new_excel_number(0)};
	// ExcelValue error_array1[] = { new_excel_number(10)}; // Not implemented
	ExcelValue error_array2[] = { TRUE, NA};
	assert(excel_and(2,true_array1).number == 1);
	assert(excel_and(1,true_array2).number == 1);
	assert(excel_and(2,false_array1).number == 0);
	assert(excel_and(2,false_array2).number == 0);
	// assert(excel_and(1,error_array1).type == ExcelError); // Not implemented
	assert(excel_and(2,error_array2).type == ExcelError);
	
	// Test AVERAGE
	ExcelValue array1[] = { new_excel_number(10), new_excel_number(5), TRUE, FALSE};
	ExcelValue array1_v = new_excel_range(array1,2,2);
	ExcelValue array2[] = { array1_v, new_excel_number(9), new_excel_string("Hello")};
	ExcelValue array3[] = { array1_v, new_excel_number(9), new_excel_string("Hello"), VALUE};
	assert(average(4, array1).number == 7.5);
	assert(average(3, array2).number == 8);
	assert(average(4, array3).type == ExcelError);
	
	// Test CHOOSE
	assert(choose(ONE,4,array1).number == 10);
	assert(choose(new_excel_number(4),4,array1).type == ExcelBoolean);
	assert(choose(new_excel_number(0),4,array1).type == ExcelError);
	assert(choose(new_excel_number(5),4,array1).type == ExcelError);
	assert(choose(ONE,4,array3).type == ExcelError);	
	
	// Test COUNT
	assert(count(4,array1).number == 2);
	assert(count(3,array2).number == 3);
	assert(count(4,array3).number == 3);
	
	// Test COUNTA
	ExcelValue count_a_test_array_1[] = { new_excel_number(10), new_excel_number(5), TRUE, FALSE, new_excel_string("Hello"), VALUE, BLANK};
  ExcelValue count_a_test_array_1_v = new_excel_range(count_a_test_array_1,7,1);
  ExcelValue count_a_test_array_2[] = {new_excel_string("Bye"),count_a_test_array_1_v};
	assert(counta(7, count_a_test_array_1).number == 6);
  assert(counta(2, count_a_test_array_2).number == 7);
	
	// Test divide
	assert(divide(new_excel_number(12.4),new_excel_number(3.2)).number == 3.875);
	assert(divide(new_excel_number(12.4),new_excel_number(0)).type == ExcelError);
	
	// Test excel_equal
	assert(excel_equal(new_excel_number(1.2),new_excel_number(3.4)).type == ExcelBoolean);
	assert(excel_equal(new_excel_number(1.2),new_excel_number(3.4)).number == false);
	assert(excel_equal(new_excel_number(1.2),new_excel_number(1.2)).number == true);
	assert(excel_equal(new_excel_string("hello"), new_excel_string("HELLO")).number == true);
	assert(excel_equal(new_excel_string("hello world"), new_excel_string("HELLO")).number == false);
	assert(excel_equal(new_excel_string("1"), ONE).number == false);
	assert(excel_equal(DIV0, ONE).type == ExcelError);

	// Test not_equal
	assert(not_equal(new_excel_number(1.2),new_excel_number(3.4)).type == ExcelBoolean);
	assert(not_equal(new_excel_number(1.2),new_excel_number(3.4)).number == true);
	assert(not_equal(new_excel_number(1.2),new_excel_number(1.2)).number == false);
	assert(not_equal(new_excel_string("hello"), new_excel_string("HELLO")).number == false);
	assert(not_equal(new_excel_string("hello world"), new_excel_string("HELLO")).number == true);
	assert(not_equal(new_excel_string("1"), ONE).number == true);
	assert(not_equal(DIV0, ONE).type == ExcelError);
	
	// Test excel_if
	// Two argument version
	assert(excel_if_2(TRUE,new_excel_number(10)).type == ExcelNumber);
	assert(excel_if_2(TRUE,new_excel_number(10)).number == 10);
	assert(excel_if_2(FALSE,new_excel_number(10)).type == ExcelBoolean);
	assert(excel_if_2(FALSE,new_excel_number(10)).number == false);
	assert(excel_if_2(NA,new_excel_number(10)).type == ExcelError);
	// Three argument version
	assert(excel_if(TRUE,new_excel_number(10),new_excel_number(20)).type == ExcelNumber);
	assert(excel_if(TRUE,new_excel_number(10),new_excel_number(20)).number == 10);
	assert(excel_if(FALSE,new_excel_number(10),new_excel_number(20)).type == ExcelNumber);
	assert(excel_if(FALSE,new_excel_number(10),new_excel_number(20)).number == 20);
	assert(excel_if(NA,new_excel_number(10),new_excel_number(20)).type == ExcelError);
	assert(excel_if(TRUE,new_excel_number(10),NA).type == ExcelNumber);
	assert(excel_if(TRUE,new_excel_number(10),NA).number == 10);
	
	// Test excel_match
	ExcelValue excel_match_array_1[] = { new_excel_number(10), new_excel_number(100) };
	ExcelValue excel_match_array_1_v = new_excel_range(excel_match_array_1,1,2);
	ExcelValue excel_match_array_2[] = { new_excel_string("Pear"), new_excel_string("Bear"), new_excel_string("Apple") };
	ExcelValue excel_match_array_2_v = new_excel_range(excel_match_array_2,3,1);
	ExcelValue excel_match_array_4[] = { ONE, BLANK, new_excel_number(0) };
	ExcelValue excel_match_array_4_v = new_excel_range(excel_match_array_4,1,3);
	ExcelValue excel_match_array_5[] = { ONE, new_excel_number(0), BLANK };
	ExcelValue excel_match_array_5_v = new_excel_range(excel_match_array_5,1,3);
	
	// Two argument version
	assert(excel_match_2(new_excel_number(10),excel_match_array_1_v).number == 1);
	assert(excel_match_2(new_excel_number(100),excel_match_array_1_v).number == 2);
	assert(excel_match_2(new_excel_number(1000),excel_match_array_1_v).type == ExcelError);
    assert(excel_match_2(new_excel_number(0), excel_match_array_4_v).number == 2);
    assert(excel_match_2(BLANK, excel_match_array_5_v).number == 2);

	// Three argument version	
    assert(excel_match(new_excel_number(10.0), excel_match_array_1_v, new_excel_number(0) ).number == 1);
    assert(excel_match(new_excel_number(100.0), excel_match_array_1_v, new_excel_number(0) ).number == 2);
    assert(excel_match(new_excel_number(1000.0), excel_match_array_1_v, new_excel_number(0) ).type == ExcelError);
    assert(excel_match(new_excel_string("bEAr"), excel_match_array_2_v, new_excel_number(0) ).number == 2);
    assert(excel_match(new_excel_number(1000.0), excel_match_array_1_v, ONE ).number == 2);
    assert(excel_match(new_excel_number(1.0), excel_match_array_1_v, ONE ).type == ExcelError);
    assert(excel_match(new_excel_string("Care"), excel_match_array_2_v, new_excel_number(-1) ).number == 1  );
    assert(excel_match(new_excel_string("Zebra"), excel_match_array_2_v, new_excel_number(-1) ).type == ExcelError);
    assert(excel_match(new_excel_string("a"), excel_match_array_2_v, new_excel_number(-1) ).number == 2);
	
	// When not given a range
    assert(excel_match(new_excel_number(10.0), new_excel_number(10), new_excel_number(0.0)).number == 1);
    assert(excel_match(new_excel_number(20.0), new_excel_number(10), new_excel_number(0.0)).type == ExcelError);
    assert(excel_match(new_excel_number(10.0), excel_match_array_1_v, BLANK).number == 1);

	// Test more than on
	// .. numbers
    assert(more_than(ONE,new_excel_number(2)).number == false);
    assert(more_than(ONE,ONE).number == false);
    assert(more_than(ONE,new_excel_number(0)).number == true);
	// .. booleans
    assert(more_than(FALSE,FALSE).number == false);
    assert(more_than(FALSE,TRUE).number == false);
    assert(more_than(TRUE,FALSE).number == true);
    assert(more_than(TRUE,TRUE).number == false);
	// ..strings
    assert(more_than(new_excel_string("HELLO"),new_excel_string("Ardvark")).number == true);		
    assert(more_than(new_excel_string("HELLO"),new_excel_string("world")).number == false);
    assert(more_than(new_excel_string("HELLO"),new_excel_string("hello")).number == false);
	// ..blanks
    assert(more_than(BLANK,ONE).number == false);
    assert(more_than(BLANK,new_excel_number(-1)).number == true);
    assert(more_than(ONE,BLANK).number == true);
    assert(more_than(new_excel_number(-1),BLANK).number == false);

	// Test less than on
	// .. numbers
    assert(less_than(ONE,new_excel_number(2)).number == true);
    assert(less_than(ONE,ONE).number == false);
    assert(less_than(ONE,new_excel_number(0)).number == false);
	// .. booleans
    assert(less_than(FALSE,FALSE).number == false);
    assert(less_than(FALSE,TRUE).number == true);
    assert(less_than(TRUE,FALSE).number == false);
    assert(less_than(TRUE,TRUE).number == false);
	// ..strings
    assert(less_than(new_excel_string("HELLO"),new_excel_string("Ardvark")).number == false);		
    assert(less_than(new_excel_string("HELLO"),new_excel_string("world")).number == true);
    assert(less_than(new_excel_string("HELLO"),new_excel_string("hello")).number == false);
	// ..blanks
    assert(less_than(BLANK,ONE).number == true);
    assert(less_than(BLANK,new_excel_number(-1)).number == false);
    assert(less_than(ONE,BLANK).number == false);
    assert(less_than(new_excel_number(-1),BLANK).number == true);

	// Test FIND function
	// ... should find the first occurrence of one string in another, returning :value if the string doesn't match
	assert(find_2(new_excel_string("one"),new_excel_string("onetwothree")).number == 1);
	assert(find_2(new_excel_string("one"),new_excel_string("twoonethree")).number == 4);
	assert(find_2(new_excel_string("one"),new_excel_string("twoonthree")).type == ExcelError);
    // ... should find the first occurrence of one string in another after a given index, returning :value if the string doesn't match
	assert(find(new_excel_string("one"),new_excel_string("onetwothree"),ONE).number == 1);
	assert(find(new_excel_string("one"),new_excel_string("twoonethree"),new_excel_number(5)).type == ExcelError);
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_number(2)).number == 4);
    // ... should be possible for the start_num to be a string, if that string converts to a number
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_string("2")).number == 4);
    // ... should return a :value error when given anything but a number as the third argument
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_string("a")).type == ExcelError);
    // ... should return a :value error when given a third argument that is less than 1 or greater than the length of the string
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_number(0)).type == ExcelError);
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_number(-1)).type == ExcelError);
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_number(7)).type == ExcelError);
	// ... BLANK in the first argument matches any character
	assert(find_2(BLANK,new_excel_string("abcdefg")).number == 1);
	assert(find(BLANK,new_excel_string("abcdefg"),new_excel_number(4)).number == 4);
    // ... should treat BLANK in the second argument as an empty string
	assert(find_2(BLANK,BLANK).number == 1);
	assert(find_2(new_excel_string("a"),BLANK).type == ExcelError);
	// ... should return an error if any argument is an error
	assert(find(new_excel_string("one"),new_excel_string("onetwothree"),NA).type == ExcelError);
	assert(find(new_excel_string("one"),NA,ONE).type == ExcelError);
	assert(find(NA,new_excel_string("onetwothree"),ONE).type == ExcelError);
	
	// Test the IFERROR function
    assert(iferror(new_excel_string("ok"),ONE).type == ExcelString);
	assert(iferror(VALUE,ONE).type == ExcelNumber);		
	
	// Test the INDEX function
	ExcelValue index_array_1[] = { new_excel_number(10), new_excel_number(20), BLANK };
	ExcelValue index_array_1_v_column = new_excel_range(index_array_1,3,1);
	ExcelValue index_array_1_v_row = new_excel_range(index_array_1,1,3);
	ExcelValue index_array_2[] = { BLANK, ONE, new_excel_number(10), new_excel_number(11), new_excel_number(100), new_excel_number(101) };
	ExcelValue index_array_2_v = new_excel_range(index_array_2,3,2);
	// ... if given one argument should return the value at that offset in the range
	assert(excel_index_2(index_array_1_v_column,new_excel_number(2.0)).number == 20);
	assert(excel_index_2(index_array_1_v_row,new_excel_number(2.0)).number == 20);
	// ... but not if the range is not a single row or single column
	assert(excel_index_2(index_array_2_v,new_excel_number(2.0)).type == ExcelError);
    // ... it should return the value in the array at position row_number, column_number
	assert(excel_index(new_excel_number(10),ONE,ONE).number == 10);
	assert(excel_index(index_array_2_v,new_excel_number(1.0),new_excel_number(2.0)).number == 1);
	assert(excel_index(index_array_2_v,new_excel_number(2.0),new_excel_number(1.0)).number == 10);
	assert(excel_index(index_array_2_v,new_excel_number(3.0),new_excel_number(1.0)).number == 100);
	assert(excel_index(index_array_2_v,new_excel_number(3.0),new_excel_number(3.0)).type == ExcelError);
	// ... it should return ZERO not blank, if a blank cell is picked
	assert(excel_index(index_array_2_v,new_excel_number(1.0),new_excel_number(1.0)).type == ExcelNumber);
	assert(excel_index(index_array_2_v,new_excel_number(1.0),new_excel_number(1.0)).number == 0);
	assert(excel_index_2(index_array_1_v_row,new_excel_number(3.0)).type == ExcelNumber);
	assert(excel_index_2(index_array_1_v_row,new_excel_number(3.0)).number == 0);
	// ... it should return the whole row if given a zero column number
	ExcelValue index_result_1_v = excel_index(index_array_2_v,new_excel_number(1.0),new_excel_number(0.0));
	assert(index_result_1_v.type == ExcelRange);
	assert(index_result_1_v.rows == 1);
	assert(index_result_1_v.columns == 2);
	ExcelValue *index_result_1_a = index_result_1_v.array;
	assert(index_result_1_a[0].number == 0);
	assert(index_result_1_a[1].number == 1);
	// ... it should return the whole column if given a zero row number
	ExcelValue index_result_2_v = excel_index(index_array_2_v,new_excel_number(0),new_excel_number(1.0));
	assert(index_result_2_v.type == ExcelRange);
	assert(index_result_2_v.rows == 3);
	assert(index_result_2_v.columns == 1);
	ExcelValue *index_result_2_a = index_result_2_v.array;
	assert(index_result_2_a[0].number == 0);
	assert(index_result_2_a[1].number == 10);
	assert(index_result_2_a[2].number == 100);
    // ... it should return a :ref error when given arguments outside array range
	assert(excel_index_2(index_array_1_v_row,new_excel_number(-1)).type == ExcelError);
	assert(excel_index_2(index_array_1_v_row,new_excel_number(4)).type == ExcelError);
    // ... it should treat BLANK as zero if given as a required row or column number
	assert(excel_index(index_array_2_v,new_excel_number(1.0),BLANK).type == ExcelRange);
	assert(excel_index(index_array_2_v,BLANK,new_excel_number(2.0)).type == ExcelRange);
    // ... it should return an error if an argument is an error
	assert(excel_index(NA,NA,NA).type == ExcelError);
	
	// LEFT(string,[characters])
	// ... should return the left n characters from a string
    assert(strcmp(left_1(new_excel_string("ONE")).string,"O") == 0);
    assert(strcmp(left(new_excel_string("ONE"),ONE).string,"O") == 0);
    assert(strcmp(left(new_excel_string("ONE"),new_excel_number(3)).string,"ONE") == 0);
	// ... should turn numbers into strings before processing
	assert(strcmp(left(new_excel_number(1.31e12),new_excel_number(3)).string, "131") == 0);
	// ... should turn booleans into the words TRUE and FALSE before processing
    assert(strcmp(left(TRUE,new_excel_number(3)).string,"TRU") == 0);
	assert(strcmp(left(FALSE,new_excel_number(3)).string,"FAL") == 0);
	// ... should return BLANK if given BLANK for either argument
	assert(left(BLANK,new_excel_number(3)).type == ExcelEmpty);
	assert(left(new_excel_string("ONE"),BLANK).type == ExcelEmpty);
	// ... should return an error if an argument is an error
    assert(left_1(NA).type == ExcelError);
    assert(left(new_excel_string("ONE"),NA).type == ExcelError);
	
	// Test less than or equal to
	// .. numbers
    assert(less_than_or_equal(ONE,new_excel_number(2)).number == true);
    assert(less_than_or_equal(ONE,ONE).number == true);
    assert(less_than_or_equal(ONE,new_excel_number(0)).number == false);
	// .. booleans
    assert(less_than_or_equal(FALSE,FALSE).number == true);
    assert(less_than_or_equal(FALSE,TRUE).number == true);
    assert(less_than_or_equal(TRUE,FALSE).number == false);
    assert(less_than_or_equal(TRUE,TRUE).number == true);
	// ..strings
    assert(less_than_or_equal(new_excel_string("HELLO"),new_excel_string("Ardvark")).number == false);		
    assert(less_than_or_equal(new_excel_string("HELLO"),new_excel_string("world")).number == true);
    assert(less_than_or_equal(new_excel_string("HELLO"),new_excel_string("hello")).number == true);
	// ..blanks
    assert(less_than_or_equal(BLANK,ONE).number == true);
    assert(less_than_or_equal(BLANK,new_excel_number(-1)).number == false);
    assert(less_than_or_equal(ONE,BLANK).number == false);
    assert(less_than_or_equal(new_excel_number(-1),BLANK).number == true);

	// Test MAX
	assert(max(4, array1).number == 10);
	assert(max(3, array2).number == 10);
	assert(max(4, array3).type == ExcelError);

	// Test MIN
	assert(min(4, array1).number == 5);
	assert(min(3, array2).number == 5);
	assert(min(4, array3).type == ExcelError);

	// Test MOD
    // ... should return the remainder of a number
	assert(mod(new_excel_number(10), new_excel_number(3)).number == 1.0);
	assert(mod(new_excel_number(10), new_excel_number(5)).number == 0.0);
    // ... should be possible for the the arguments to be strings, if they convert to a number
	assert(mod(new_excel_string("3.5"),new_excel_string("2")).number == 1.5);
    // ... should treat BLANK as zero
	assert(mod(BLANK,new_excel_number(10)).number == 0);
	assert(mod(new_excel_number(10),BLANK).type == ExcelError);
	assert(mod(BLANK,BLANK).type == ExcelError);
    // ... should treat true as 1 and FALSE as 0
	assert((mod(new_excel_number(1.1),TRUE).number - 0.1) < 0.001);	
	assert(mod(new_excel_number(1.1),FALSE).type == ExcelError);
	assert(mod(FALSE,new_excel_number(10)).number == 0);
    // ... should return an error when given inappropriate arguments
	assert(mod(new_excel_string("Asdasddf"),new_excel_string("adsfads")).type == ExcelError);
    // ... should return an error if an argument is an error
	assert(mod(new_excel_number(1),VALUE).type == ExcelError);
	assert(mod(VALUE,new_excel_number(1)).type == ExcelError);
	assert(mod(VALUE,VALUE).type == ExcelError);
	
	// Test more than or equal to on
	// .. numbers
    assert(more_than_or_equal(ONE,new_excel_number(2)).number == false);
    assert(more_than_or_equal(ONE,ONE).number == true);
    assert(more_than_or_equal(ONE,new_excel_number(0)).number == true);
	// .. booleans
    assert(more_than_or_equal(FALSE,FALSE).number == true);
    assert(more_than_or_equal(FALSE,TRUE).number == false);
    assert(more_than_or_equal(TRUE,FALSE).number == true);
    assert(more_than_or_equal(TRUE,TRUE).number == true);
	// ..strings
    assert(more_than_or_equal(new_excel_string("HELLO"),new_excel_string("Ardvark")).number == true);		
    assert(more_than_or_equal(new_excel_string("HELLO"),new_excel_string("world")).number == false);
    assert(more_than_or_equal(new_excel_string("HELLO"),new_excel_string("hello")).number == true);
	// ..blanks
    assert(more_than_or_equal(BLANK,BLANK).number == true);
    assert(more_than_or_equal(BLANK,ONE).number == false);
    assert(more_than_or_equal(BLANK,new_excel_number(-1)).number == true);
    assert(more_than_or_equal(ONE,BLANK).number == true);
    assert(more_than_or_equal(new_excel_number(-1),BLANK).number == false);	
	
	// Test negative
    // ... should return the negative of its arguments
	assert(negative(new_excel_number(1)).number == -1);
	assert(negative(new_excel_number(-1)).number == 1);
    // ... should treat strings that only contain numbers as numbers
	assert(negative(new_excel_string("10")).number == -10);
	assert(negative(new_excel_string("-1.3")).number == 1.3);
    // ... should return an error when given inappropriate arguments
	assert(negative(new_excel_string("Asdasddf")).type == ExcelError);
    // ... should treat BLANK as zero
	assert(negative(BLANK).number == 0);
	
	// Test PMT(rate,number_of_periods,present_value) - optional arguments not yet implemented
    // ... should calculate the monthly payment required for a given principal, interest rate and loan period
    assert((pmt(new_excel_number(0.1),new_excel_number(10),new_excel_number(100)).number - -16.27) < 0.01);
    assert((pmt(new_excel_number(0.0123),new_excel_number(99.1),new_excel_number(123.32)).number - -2.159) < 0.01);
    assert((pmt(new_excel_number(0),new_excel_number(2),new_excel_number(10)).number - -5) < 0.01);

	// Test power
    // ... should return sum of its arguments
	assert(power(new_excel_number(2),new_excel_number(3)).number == 8);
	assert(power(new_excel_number(4.0),new_excel_number(0.5)).number == 2.0);
	
	// Test round
    assert(excel_round(new_excel_number(1.1), new_excel_number(0)).number == 1.0);
    assert(excel_round(new_excel_number(1.5), new_excel_number(0)).number == 2.0);
    assert(excel_round(new_excel_number(1.56),new_excel_number(1)).number == 1.6);
    assert(excel_round(new_excel_number(-1.56),new_excel_number(1)).number == -1.6);

	// Test rounddown
    assert(rounddown(new_excel_number(1.1), new_excel_number(0)).number == 1.0);
    assert(rounddown(new_excel_number(1.5), new_excel_number(0)).number == 1.0);
    assert(rounddown(new_excel_number(1.56),new_excel_number(1)).number == 1.5);
    assert(rounddown(new_excel_number(-1.56),new_excel_number(1)).number == -1.5);	

	// Test roundup
    assert(roundup(new_excel_number(1.1), new_excel_number(0)).number == 2.0);
    assert(roundup(new_excel_number(1.5), new_excel_number(0)).number == 2.0);
    assert(roundup(new_excel_number(1.56),new_excel_number(1)).number == 1.6);
    assert(roundup(new_excel_number(-1.56),new_excel_number(1)).number == -1.6);	
	
	// Test string joining
	ExcelValue string_join_array_1[] = {new_excel_string("Hello "), new_excel_string("world")};
	ExcelValue string_join_array_2[] = {new_excel_string("Hello "), new_excel_string("world"), new_excel_string("!")};
	ExcelValue string_join_array_3[] = {new_excel_string("Top "), new_excel_number(10.0)};
	ExcelValue string_join_array_4[] = {new_excel_string("Top "), new_excel_number(10.5)};	
	ExcelValue string_join_array_5[] = {new_excel_string("Top "), TRUE, FALSE};	
	// ... should return a string by combining its arguments
	// inspect_excel_value(string_join(2, string_join_array_1));
  assert(string_join(2, string_join_array_1).string[6] == 'w');
  assert(string_join(2, string_join_array_1).string[11] == '\0');
	// ... should cope with an arbitrary number of arguments
  assert(string_join(3, string_join_array_2).string[11] == '!');
  assert(string_join(3, string_join_array_3).string[12] == '\0');
	// ... should convert values to strings as it goes
  assert(string_join(2, string_join_array_3).string[4] == '1');
  assert(string_join(2, string_join_array_3).string[5] == '0');
  assert(string_join(2, string_join_array_3).string[6] == '\0');
	// ... should convert integer values into strings without decimal points, and float values with decimal points
  assert(string_join(2, string_join_array_4).string[4] == '1');
  assert(string_join(2, string_join_array_4).string[5] == '0');
  assert(string_join(2, string_join_array_4).string[6] == '.');
  assert(string_join(2, string_join_array_4).string[7] == '5');
  assert(string_join(2, string_join_array_4).string[8] == '\0');
	// ... should convert TRUE and FALSE into strings
  assert(string_join(3,string_join_array_5).string[4] == 'T');
	
  // Test SUBTOTAL function
  ExcelValue subtotal_array_1[] = {new_excel_number(10),new_excel_number(100),BLANK};
  ExcelValue subtotal_array_1_v = new_excel_range(subtotal_array_1,3,1);
  ExcelValue subtotal_array_2[] = {new_excel_number(1),new_excel_string("two"),subtotal_array_1_v};
  
  // new_excel_number(1.0); 
  // inspect_excel_value(new_excel_number(1.0)); 
  // inspect_excel_value(new_excel_range(subtotal_array_2,3,1)); 
  // inspect_excel_value(subtotal(new_excel_number(1.0),3,subtotal_array_2)); 
  
  assert(subtotal(new_excel_number(1.0),3,subtotal_array_2).number == 111.0/3.0);
  assert(subtotal(new_excel_number(2.0),3,subtotal_array_2).number == 3);
  assert(subtotal(new_excel_number(3.0),7, count_a_test_array_1).number == 6);
  assert(subtotal(new_excel_number(3.0),3,subtotal_array_2).number == 4);
  assert(subtotal(new_excel_number(9.0),3,subtotal_array_2).number == 111);
  assert(subtotal(new_excel_number(101.0),3,subtotal_array_2).number == 111.0/3.0);
  assert(subtotal(new_excel_number(102.0),3,subtotal_array_2).number == 3);
  assert(subtotal(new_excel_number(103.0),3,subtotal_array_2).number == 4);
  assert(subtotal(new_excel_number(109.0),3,subtotal_array_2).number == 111);
  
  // Test SUMIFS function
  ExcelValue sumifs_array_1[] = {new_excel_number(10),new_excel_number(100),BLANK};
  ExcelValue sumifs_array_1_v = new_excel_range(sumifs_array_1,3,1);
  ExcelValue sumifs_array_2[] = {new_excel_string("pear"),new_excel_string("bear"),new_excel_string("apple")};
  ExcelValue sumifs_array_2_v = new_excel_range(sumifs_array_2,3,1);
  ExcelValue sumifs_array_3[] = {new_excel_number(1),new_excel_number(2),new_excel_number(3),new_excel_number(4),new_excel_number(5),new_excel_number(5)};
  ExcelValue sumifs_array_3_v = new_excel_range(sumifs_array_3,6,1);
  ExcelValue sumifs_array_4[] = {new_excel_string("CO2"),new_excel_string("CH4"),new_excel_string("N2O"),new_excel_string("CH4"),new_excel_string("N2O"),new_excel_string("CO2")};
  ExcelValue sumifs_array_4_v = new_excel_range(sumifs_array_4,6,1);
  ExcelValue sumifs_array_5[] = {new_excel_string("1A"),new_excel_string("1A"),new_excel_string("1A"),new_excel_number(4),new_excel_number(4),new_excel_number(5)};
  ExcelValue sumifs_array_5_v = new_excel_range(sumifs_array_5,6,1);
  
  // ... should only sum values that meet all of the criteria
  ExcelValue sumifs_array_6[] = { sumifs_array_1_v, new_excel_number(10), sumifs_array_2_v, new_excel_string("Bear") };
  assert(sumifs(sumifs_array_1_v,4,sumifs_array_6).number == 0.0);
  
  ExcelValue sumifs_array_7[] = { sumifs_array_1_v, new_excel_number(10), sumifs_array_2_v, new_excel_string("Pear") };
  assert(sumifs(sumifs_array_1_v,4,sumifs_array_7).number == 10.0);
  
  // ... should work when single cells are given where ranges expected
  ExcelValue sumifs_array_8[] = { new_excel_string("CAR"), new_excel_string("CAR"), new_excel_string("FCV"), new_excel_string("FCV")};
  assert(sumifs(new_excel_number(0.143897265452564), 4, sumifs_array_8).number == 0.143897265452564);

  // ... should match numbers with strings that contain numbers
  ExcelValue sumifs_array_9[] = { new_excel_number(10), new_excel_string("10.0")};
  assert(sumifs(new_excel_number(100),2,sumifs_array_9).number == 100);
  
  ExcelValue sumifs_array_10[] = { sumifs_array_4_v, new_excel_string("CO2"), sumifs_array_5_v, new_excel_number(2)};
  assert(sumifs(sumifs_array_3_v,4, sumifs_array_10).number == 0);
  
  // ... should match with strings that contain criteria
  ExcelValue sumifs_array_10a[] = { sumifs_array_3_v, new_excel_string("=5")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10a).number == 10);

  ExcelValue sumifs_array_10b[] = { sumifs_array_3_v, new_excel_string("<>3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10b).number == 17);

  ExcelValue sumifs_array_10c[] = { sumifs_array_3_v, new_excel_string("<3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10c).number == 3);
  
  ExcelValue sumifs_array_10d[] = { sumifs_array_3_v, new_excel_string("<=3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10d).number == 6);

  ExcelValue sumifs_array_10e[] = { sumifs_array_3_v, new_excel_string(">3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10e).number == 14);

  ExcelValue sumifs_array_10f[] = { sumifs_array_3_v, new_excel_string(">=3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10f).number == 17);
  
  // ... should treat BLANK as an empty string when in the check_range, but not in the criteria
  ExcelValue sumifs_array_11[] = { BLANK, new_excel_number(20)};
  assert(sumifs(new_excel_number(100),2,sumifs_array_11).number == 0);
  
  ExcelValue sumifs_array_12[] = {BLANK, new_excel_string("")};
  assert(sumifs(new_excel_number(100),2,sumifs_array_12).number == 100);
  
  ExcelValue sumifs_array_13[] = {BLANK, BLANK};
  assert(sumifs(new_excel_number(100),2,sumifs_array_13).number == 0);
    
  // ... should return an error if range argument is an error
  assert(sumifs(REF,2,sumifs_array_13).type == ExcelError);
  
  
  // Test SUMIF
  // ... where there is only a check range
  assert(sumif_2(sumifs_array_1_v,new_excel_string(">0")).number == 110.0);
  assert(sumif_2(sumifs_array_1_v,new_excel_string(">10")).number == 100.0);
  assert(sumif_2(sumifs_array_1_v,new_excel_string("<100")).number == 10.0);
  
  // ... where there is a seprate sum range
  ExcelValue sumif_array_1[] = {new_excel_number(15),new_excel_number(20), new_excel_number(30)};
  ExcelValue sumif_array_1_v = new_excel_range(sumif_array_1,3,1);
  assert(sumif(sumifs_array_1_v,new_excel_string("10"),sumif_array_1_v).number == 15);
  
  
  // Test SUMPRODUCT
  ExcelValue sumproduct_1[] = { new_excel_number(10), new_excel_number(100), BLANK};
  ExcelValue sumproduct_2[] = { BLANK, new_excel_number(100), new_excel_number(10), BLANK};
  ExcelValue sumproduct_3[] = { BLANK };
  ExcelValue sumproduct_4[] = { new_excel_number(10), new_excel_number(100), new_excel_number(1000)};
  ExcelValue sumproduct_5[] = { new_excel_number(1), new_excel_number(2), new_excel_number(3)};
  ExcelValue sumproduct_6[] = { new_excel_number(1), new_excel_number(2), new_excel_number(4), new_excel_number(5)};
  ExcelValue sumproduct_7[] = { new_excel_number(10), new_excel_number(20), new_excel_number(40), new_excel_number(50)};
  ExcelValue sumproduct_8[] = { new_excel_number(11), new_excel_number(21), new_excel_number(41), new_excel_number(51)};
  ExcelValue sumproduct_9[] = { BLANK, BLANK };
  
  ExcelValue sumproduct_1_v = new_excel_range( sumproduct_1, 3, 1);
  ExcelValue sumproduct_2_v = new_excel_range( sumproduct_2, 3, 1);
  ExcelValue sumproduct_3_v = new_excel_range( sumproduct_3, 1, 1);
  // ExcelValue sumproduct_4_v = new_excel_range( sumproduct_4, 1, 3); // Unused
  ExcelValue sumproduct_5_v = new_excel_range( sumproduct_5, 3, 1);
  ExcelValue sumproduct_6_v = new_excel_range( sumproduct_6, 2, 2);
  ExcelValue sumproduct_7_v = new_excel_range( sumproduct_7, 2, 2);
  ExcelValue sumproduct_8_v = new_excel_range( sumproduct_8, 2, 2);
  ExcelValue sumproduct_9_v = new_excel_range( sumproduct_9, 2, 1);
  
  // ... should multiply together and then sum the elements in row or column areas given as arguments
  ExcelValue sumproducta_1[] = {sumproduct_1_v, sumproduct_2_v};
  assert(sumproduct(2,sumproducta_1).number == 100*100);

  // ... should return :value when miss-matched array sizes
  ExcelValue sumproducta_2[] = {sumproduct_1_v, sumproduct_3_v};
  assert(sumproduct(2,sumproducta_2).type == ExcelError);

  // ... if all its arguments are single values, should multiply them together
  // ExcelValue *sumproducta_3 = sumproduct_4;
  assert(sumproduct(3,sumproduct_4).number == 10*100*1000);

  // ... if it only has one range as an argument, should add its elements together
  ExcelValue sumproducta_4[] = {sumproduct_5_v};
  assert(sumproduct(1,sumproducta_4).number == 1 + 2 + 3);

  // ... if given multi row and column areas as arguments, should multipy the corresponding cell in each area and then add them all
  ExcelValue sumproducta_5[] = {sumproduct_6_v, sumproduct_7_v, sumproduct_8_v};
  assert(sumproduct(3,sumproducta_5).number == 1*10*11 + 2*20*21 + 4*40*41 + 5*50*51);

  // ... should raise an error if BLANK values outside of an array
  ExcelValue sumproducta_6[] = {BLANK,new_excel_number(1)};
  assert(sumproduct(2,sumproducta_6).type == ExcelError);

  // ... should ignore non-numeric values within an array
  ExcelValue sumproducta_7[] = {sumproduct_9_v, sumproduct_9_v};
  assert(sumproduct(2,sumproducta_7).number == 0);

  // ... should return an error if an argument is an error
  ExcelValue sumproducta_8[] = {VALUE};
  assert(sumproduct(1,sumproducta_8).type == ExcelError);
  
  // Test VLOOKUP
  ExcelValue vlookup_a1[] = {new_excel_number(1),new_excel_number(10),new_excel_number(2),new_excel_number(20),new_excel_number(3),new_excel_number(30)};
  ExcelValue vlookup_a2[] = {new_excel_string("hello"),new_excel_number(10),new_excel_number(2),new_excel_number(20),new_excel_number(3),new_excel_number(30)};
  ExcelValue vlookup_a3[] = {BLANK,new_excel_number(10),new_excel_number(2),new_excel_number(20),new_excel_number(3),new_excel_number(30)};
  ExcelValue vlookup_a1_v = new_excel_range(vlookup_a1,3,2);
  ExcelValue vlookup_a2_v = new_excel_range(vlookup_a2,3,2);
  ExcelValue vlookup_a3_v = new_excel_range(vlookup_a3,3,2);
  // ... should match the first argument against the first column of the table in the second argument, returning the value in the column specified by the third argument
  assert(vlookup_3(new_excel_number(2.0),vlookup_a1_v,new_excel_number(2)).number == 20);
  assert(vlookup_3(new_excel_number(1.5),vlookup_a1_v,new_excel_number(2)).number == 10);
  assert(vlookup_3(new_excel_number(0.5),vlookup_a1_v,new_excel_number(2)).type == ExcelError);
  assert(vlookup_3(new_excel_number(10),vlookup_a1_v,new_excel_number(2)).number == 30);
  assert(vlookup_3(new_excel_number(2.6),vlookup_a1_v,new_excel_number(2)).number == 20);
  // ... has a four argument variant that matches the lookup type
  assert(vlookup(new_excel_number(2.6),vlookup_a1_v,new_excel_number(2),TRUE).number == 20);
  assert(vlookup(new_excel_number(2.6),vlookup_a1_v,new_excel_number(2),FALSE).type == ExcelError);
  assert(vlookup(new_excel_string("HELLO"),vlookup_a2_v,new_excel_number(2),FALSE).number == 10);
  assert(vlookup(new_excel_string("HELMP"),vlookup_a2_v,new_excel_number(2),TRUE).number == 10);
  // ... BLANK should not match with anything" do
  assert(vlookup_3(BLANK,vlookup_a3_v,new_excel_number(2)).type == ExcelError);
  // ... should return an error if an argument is an error" do
  assert(vlookup(VALUE,vlookup_a1_v,new_excel_number(2),FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),VALUE,new_excel_number(2),FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),vlookup_a1_v,VALUE,FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),vlookup_a1_v,new_excel_number(2),VALUE).type == ExcelError);
  assert(vlookup(VALUE,VALUE,VALUE,VALUE).type == ExcelError);
	
  // Test SUM
  ExcelValue sum_array_0[] = {new_excel_number(1084.4557258064517),new_excel_number(32.0516914516129),new_excel_number(137.36439193548387)};
  ExcelValue sum_array_0_v = new_excel_range(sum_array_0,3,1);
  ExcelValue sum_array_1[] = {sum_array_0_v};
  assert(sum(1,sum_array_1).number == 1253.8718091935484);
  
  // Release memory
  free_all_allocated_memory();
  
  return 0;
}

int main() {
	return test_functions();
}
// End of the generic c functions

// Start of the file specific functions

// definitions
static ExcelValue _common0();
static ExcelValue _common1();
static ExcelValue _common2();
static ExcelValue _common3();
static ExcelValue _common4();
static ExcelValue _common5();
static ExcelValue _common6();
static ExcelValue _common7();
static ExcelValue _common8();
static ExcelValue _common9();
static ExcelValue _common10();
static ExcelValue _common11();
static ExcelValue _common12();
static ExcelValue _common13();
static ExcelValue _common14();
static ExcelValue _common15();
static ExcelValue _common16();
static ExcelValue _common17();
static ExcelValue _common18();
static ExcelValue _common19();
static ExcelValue _common20();
static ExcelValue _common21();
static ExcelValue _common22();
static ExcelValue _common23();
static ExcelValue _common24();
static ExcelValue _common25();
static ExcelValue _common26();
static ExcelValue _common27();
static ExcelValue _common28();
static ExcelValue _common29();
static ExcelValue _common30();
static ExcelValue _common31();
static ExcelValue _common32();
static ExcelValue _common33();
static ExcelValue _common34();
static ExcelValue _common35();
static ExcelValue _common36();
static ExcelValue _common37();
static ExcelValue _common38();
static ExcelValue _common39();
static ExcelValue _common40();
static ExcelValue _common41();
static ExcelValue _common42();
static ExcelValue _common43();
static ExcelValue _common44();
static ExcelValue _common45();
static ExcelValue _common46();
static ExcelValue _common47();
static ExcelValue _common48();
static ExcelValue _common49();
static ExcelValue _common50();
static ExcelValue _common51();
static ExcelValue _common52();
static ExcelValue _common53();
static ExcelValue _common54();
static ExcelValue _common55();
static ExcelValue _common56();
static ExcelValue _common57();
static ExcelValue _common58();
static ExcelValue _common59();
static ExcelValue _common60();
static ExcelValue _common61();
static ExcelValue _common62();
static ExcelValue _common63();
static ExcelValue _common64();
static ExcelValue _common65();
static ExcelValue _common66();
static ExcelValue _common67();
static ExcelValue _common68();
static ExcelValue _common69();
static ExcelValue _common70();
static ExcelValue _common71();
static ExcelValue _common72();
static ExcelValue _common73();
static ExcelValue _common74();
static ExcelValue _common75();
static ExcelValue _common76();
static ExcelValue _common77();
static ExcelValue _common78();
static ExcelValue _common79();
static ExcelValue _common80();
static ExcelValue _common81();
static ExcelValue _common82();
static ExcelValue _common83();
static ExcelValue _common84();
static ExcelValue _common85();
static ExcelValue _common86();
static ExcelValue _common87();
static ExcelValue _common88();
static ExcelValue _common89();
static ExcelValue _common90();
static ExcelValue _common91();
static ExcelValue _common92();
static ExcelValue _common93();
static ExcelValue _common94();
static ExcelValue _common95();
static ExcelValue _common96();
static ExcelValue _common97();
static ExcelValue _common98();
static ExcelValue _common99();
static ExcelValue _common100();
static ExcelValue _common101();
static ExcelValue _common102();
static ExcelValue _common103();
static ExcelValue _common104();
static ExcelValue _common105();
static ExcelValue _common106();
static ExcelValue _common107();
static ExcelValue _common108();
static ExcelValue _common109();
static ExcelValue _common110();
static ExcelValue _common111();
static ExcelValue _common112();
static ExcelValue _common113();
static ExcelValue _common114();
static ExcelValue _common115();
static ExcelValue _common116();
static ExcelValue _common117();
static ExcelValue _common118();
static ExcelValue _common119();
static ExcelValue _common120();
static ExcelValue _common121();
static ExcelValue _common122();
static ExcelValue _common123();
static ExcelValue _common124();
static ExcelValue _common125();
static ExcelValue _common126();
static ExcelValue _common127();
static ExcelValue _common128();
static ExcelValue _common129();
static ExcelValue _common130();
static ExcelValue _common131();
static ExcelValue _common132();
static ExcelValue _common133();
static ExcelValue _common134();
static ExcelValue _common135();
static ExcelValue _common136();
static ExcelValue _common137();
static ExcelValue _common138();
static ExcelValue _common139();
static ExcelValue _common140();
static ExcelValue _common141();
static ExcelValue _common142();
static ExcelValue _common143();
static ExcelValue _common144();
static ExcelValue _common145();
static ExcelValue _common146();
static ExcelValue _common147();
static ExcelValue _common148();
static ExcelValue _common149();
static ExcelValue _common150();
static ExcelValue _common151();
static ExcelValue _common152();
static ExcelValue _common153();
static ExcelValue _common154();
static ExcelValue _common155();
static ExcelValue _common156();
static ExcelValue _common157();
static ExcelValue _common158();
static ExcelValue _common159();
static ExcelValue _common160();
static ExcelValue _common161();
static ExcelValue _common162();
static ExcelValue _common163();
static ExcelValue _common164();
static ExcelValue _common165();
static ExcelValue _common166();
static ExcelValue _common167();
static ExcelValue _common168();
static ExcelValue _common169();
static ExcelValue _common170();
static ExcelValue _common171();
static ExcelValue _common172();
static ExcelValue _common173();
static ExcelValue _common174();
static ExcelValue _common175();
static ExcelValue _common176();
static ExcelValue _common177();
static ExcelValue _common178();
static ExcelValue _common179();
static ExcelValue _common180();
static ExcelValue _common181();
static ExcelValue _common182();
static ExcelValue _common183();
static ExcelValue _common184();
static ExcelValue _common185();
static ExcelValue _common186();
static ExcelValue _common187();
static ExcelValue _common188();
static ExcelValue _common189();
static ExcelValue _common190();
static ExcelValue _common191();
static ExcelValue _common192();
static ExcelValue _common193();
static ExcelValue _common194();
static ExcelValue _common195();
static ExcelValue _common196();
static ExcelValue _common197();
static ExcelValue _common198();
static ExcelValue _common199();
static ExcelValue _common200();
static ExcelValue _common201();
static ExcelValue _common202();
static ExcelValue _common203();
static ExcelValue _common204();
static ExcelValue _common205();
static ExcelValue _common206();
static ExcelValue _common207();
static ExcelValue _common208();
static ExcelValue _common209();
static ExcelValue _common210();
static ExcelValue _common211();
static ExcelValue _common212();
static ExcelValue _common213();
static ExcelValue _common214();
static ExcelValue _common215();
static ExcelValue _common216();
static ExcelValue _common217();
static ExcelValue _common218();
static ExcelValue _common219();
static ExcelValue _common220();
static ExcelValue _common221();
static ExcelValue _common222();
static ExcelValue _common223();
static ExcelValue _common224();
static ExcelValue _common225();
static ExcelValue _common226();
static ExcelValue _common227();
static ExcelValue _common228();
static ExcelValue _common229();
static ExcelValue _common230();
static ExcelValue _common231();
static ExcelValue _common232();
static ExcelValue _common233();
static ExcelValue _common234();
static ExcelValue _common235();
static ExcelValue _common236();
static ExcelValue _common237();
static ExcelValue _common238();
static ExcelValue _common239();
static ExcelValue _common240();
static ExcelValue _common241();
static ExcelValue _common242();
static ExcelValue _common243();
static ExcelValue _common244();
static ExcelValue _common245();
static ExcelValue _common246();
static ExcelValue _common247();
static ExcelValue _common248();
static ExcelValue _common249();
static ExcelValue _common250();
static ExcelValue _common251();
static ExcelValue _common252();
static ExcelValue _common253();
static ExcelValue _common254();
static ExcelValue _common255();
static ExcelValue _common256();
static ExcelValue _common257();
static ExcelValue _common258();
static ExcelValue _common259();
static ExcelValue _common260();
static ExcelValue _common261();
static ExcelValue _common262();
static ExcelValue _common263();
static ExcelValue _common264();
static ExcelValue _common265();
static ExcelValue _common266();
static ExcelValue _common267();
static ExcelValue _common268();
static ExcelValue _common269();
static ExcelValue _common270();
static ExcelValue _common271();
static ExcelValue _common272();
static ExcelValue _common273();
static ExcelValue _common274();
static ExcelValue _common275();
static ExcelValue _common276();
static ExcelValue _common277();
static ExcelValue _common278();
static ExcelValue _common279();
static ExcelValue _common280();
static ExcelValue _common281();
static ExcelValue _common282();
static ExcelValue _common283();
static ExcelValue _common284();
static ExcelValue _common285();
static ExcelValue _common286();
static ExcelValue _common287();
static ExcelValue _common288();
static ExcelValue _common289();
static ExcelValue _common290();
static ExcelValue _common291();
static ExcelValue _common292();
static ExcelValue _common293();
static ExcelValue _common294();
static ExcelValue _common295();
static ExcelValue _common296();
static ExcelValue _common297();
static ExcelValue _common298();
static ExcelValue _common299();
static ExcelValue _common300();
static ExcelValue _common301();
static ExcelValue _common302();
static ExcelValue _common303();
static ExcelValue _common304();
static ExcelValue _common305();
static ExcelValue _common306();
static ExcelValue _common307();
static ExcelValue _common308();
static ExcelValue _common309();
static ExcelValue _common310();
static ExcelValue _common311();
static ExcelValue _common312();
static ExcelValue _common313();
static ExcelValue _common314();
static ExcelValue _common315();
static ExcelValue _common316();
static ExcelValue _common317();
static ExcelValue _common318();
static ExcelValue _common319();
static ExcelValue _common320();
static ExcelValue _common321();
static ExcelValue _common322();
static ExcelValue _common323();
static ExcelValue _common324();
static ExcelValue _common325();
static ExcelValue _common326();
static ExcelValue _common327();
static ExcelValue _common328();
static ExcelValue _common329();
static ExcelValue _common330();
static ExcelValue _common331();
static ExcelValue _common332();
static ExcelValue _common333();
static ExcelValue _common334();
static ExcelValue _common335();
static ExcelValue _common336();
static ExcelValue _common337();
static ExcelValue _common338();
static ExcelValue _common339();
static ExcelValue _common340();
static ExcelValue _common341();
static ExcelValue _common342();
static ExcelValue _common343();
static ExcelValue _common344();
static ExcelValue _common345();
static ExcelValue _common346();
static ExcelValue _common347();
static ExcelValue _common348();
static ExcelValue _common349();
static ExcelValue _common350();
static ExcelValue _common351();
static ExcelValue _common352();
static ExcelValue _common353();
static ExcelValue _common354();
static ExcelValue _common355();
static ExcelValue _common356();
static ExcelValue _common357();
static ExcelValue _common358();
static ExcelValue _common359();
static ExcelValue _common360();
static ExcelValue _common361();
static ExcelValue _common362();
static ExcelValue _common363();
static ExcelValue _common364();
static ExcelValue _common365();
static ExcelValue _common366();
static ExcelValue _common367();
static ExcelValue _common368();
static ExcelValue _common369();
static ExcelValue _common370();
static ExcelValue _common371();
static ExcelValue _common372();
static ExcelValue _common373();
static ExcelValue _common374();
static ExcelValue _common375();
static ExcelValue _common376();
static ExcelValue _common377();
static ExcelValue _common378();
static ExcelValue _common379();
static ExcelValue _common380();
static ExcelValue _common381();
static ExcelValue _common382();
static ExcelValue _common383();
static ExcelValue _common384();
static ExcelValue _common385();
static ExcelValue _common386();
static ExcelValue _common387();
static ExcelValue _common388();
static ExcelValue _common389();
static ExcelValue _common390();
static ExcelValue _common391();
static ExcelValue _common392();
static ExcelValue _common393();
static ExcelValue _common394();
static ExcelValue _common395();
static ExcelValue _common396();
static ExcelValue _common397();
static ExcelValue _common398();
static ExcelValue _common399();
static ExcelValue _common400();
static ExcelValue _common401();
static ExcelValue _common402();
static ExcelValue _common403();
static ExcelValue _common404();
static ExcelValue _common405();
static ExcelValue _common406();
static ExcelValue _common407();
static ExcelValue _common408();
static ExcelValue _common409();
static ExcelValue _common410();
static ExcelValue _common411();
static ExcelValue _common412();
static ExcelValue _common413();
static ExcelValue _common414();
static ExcelValue _common415();
static ExcelValue _common416();
static ExcelValue _common417();
static ExcelValue _common418();
static ExcelValue _common419();
static ExcelValue _common420();
static ExcelValue _common421();
static ExcelValue _common422();
static ExcelValue _common423();
static ExcelValue _common424();
static ExcelValue _common425();
static ExcelValue _common426();
static ExcelValue _common427();
static ExcelValue _common428();
static ExcelValue _common429();
static ExcelValue _common430();
static ExcelValue _common431();
static ExcelValue _common432();
static ExcelValue _common433();
static ExcelValue _common434();
ExcelValue model_b3();
ExcelValue model_f3();
ExcelValue model_b4();
ExcelValue model_f6();
ExcelValue model_b7();
ExcelValue model_f7();
ExcelValue model_b8();
ExcelValue model_b9();
ExcelValue model_b10();
ExcelValue model_b11();
ExcelValue model_b12();
ExcelValue model_b13();
ExcelValue model_b31();
ExcelValue model_b32();
ExcelValue model_b34();
ExcelValue model_b35();
ExcelValue model_b36();
ExcelValue model_b37();
static ExcelValue model_n38();
ExcelValue model_b40();
ExcelValue model_c40();
ExcelValue model_d40();
ExcelValue model_b44();
ExcelValue model_c44();
ExcelValue model_b45();
ExcelValue model_c45();
static ExcelValue model_f47();
static ExcelValue model_g47();
static ExcelValue model_h47();
static ExcelValue model_i47();
static ExcelValue model_j47();
static ExcelValue model_k47();
static ExcelValue model_l47();
static ExcelValue model_m47();
static ExcelValue model_n47();
static ExcelValue model_o47();
static ExcelValue model_p47();
static ExcelValue model_q47();
static ExcelValue model_r47();
static ExcelValue model_s47();
static ExcelValue model_t47();
static ExcelValue model_u47();
static ExcelValue model_v47();
static ExcelValue model_w47();
static ExcelValue model_x47();
static ExcelValue model_y47();
static ExcelValue model_z47();
static ExcelValue model_aa47();
static ExcelValue model_ab47();
static ExcelValue model_ac47();
static ExcelValue model_ad47();
static ExcelValue model_ae47();
static ExcelValue model_af47();
static ExcelValue model_ag47();
static ExcelValue model_ah47();
static ExcelValue model_ai47();
static ExcelValue model_aj47();
static ExcelValue model_ak47();
static ExcelValue model_al47();
static ExcelValue model_am47();
static ExcelValue model_an47();
ExcelValue model_b48();
ExcelValue model_c48();
ExcelValue model_d48();
ExcelValue model_e48();
ExcelValue model_f48();
ExcelValue model_g48();
ExcelValue model_h48();
ExcelValue model_i48();
ExcelValue model_j48();
ExcelValue model_k48();
ExcelValue model_l48();
ExcelValue model_m48();
ExcelValue model_n48();
ExcelValue model_o48();
ExcelValue model_p48();
ExcelValue model_q48();
ExcelValue model_r48();
ExcelValue model_s48();
ExcelValue model_t48();
ExcelValue model_u48();
ExcelValue model_v48();
ExcelValue model_w48();
ExcelValue model_x48();
ExcelValue model_y48();
ExcelValue model_z48();
ExcelValue model_aa48();
ExcelValue model_ab48();
ExcelValue model_ac48();
ExcelValue model_ad48();
ExcelValue model_ae48();
ExcelValue model_af48();
ExcelValue model_ag48();
ExcelValue model_ah48();
ExcelValue model_ai48();
ExcelValue model_aj48();
ExcelValue model_ak48();
ExcelValue model_al48();
ExcelValue model_am48();
ExcelValue model_an48();
ExcelValue model_b49();
ExcelValue model_c49();
ExcelValue model_d49();
ExcelValue model_e49();
ExcelValue model_f49();
ExcelValue model_g49();
ExcelValue model_h49();
ExcelValue model_i49();
ExcelValue model_j49();
ExcelValue model_k49();
ExcelValue model_l49();
ExcelValue model_m49();
ExcelValue model_n49();
ExcelValue model_o49();
ExcelValue model_p49();
ExcelValue model_q49();
ExcelValue model_r49();
ExcelValue model_s49();
ExcelValue model_t49();
ExcelValue model_u49();
ExcelValue model_v49();
ExcelValue model_w49();
ExcelValue model_x49();
ExcelValue model_y49();
ExcelValue model_z49();
ExcelValue model_aa49();
ExcelValue model_ab49();
ExcelValue model_ac49();
ExcelValue model_ad49();
ExcelValue model_ae49();
ExcelValue model_af49();
ExcelValue model_ag49();
ExcelValue model_ah49();
ExcelValue model_ai49();
ExcelValue model_aj49();
ExcelValue model_ak49();
ExcelValue model_al49();
ExcelValue model_am49();
ExcelValue model_an49();
ExcelValue model_b50();
ExcelValue model_c50();
ExcelValue model_d50();
ExcelValue model_e50();
ExcelValue model_f50();
ExcelValue model_g50();
ExcelValue model_h50();
ExcelValue model_i50();
ExcelValue model_j50();
ExcelValue model_k50();
ExcelValue model_l50();
ExcelValue model_m50();
ExcelValue model_n50();
ExcelValue model_o50();
ExcelValue model_p50();
ExcelValue model_q50();
ExcelValue model_r50();
ExcelValue model_s50();
ExcelValue model_t50();
ExcelValue model_u50();
ExcelValue model_v50();
ExcelValue model_w50();
ExcelValue model_x50();
ExcelValue model_y50();
ExcelValue model_z50();
ExcelValue model_aa50();
ExcelValue model_ab50();
ExcelValue model_ac50();
ExcelValue model_ad50();
ExcelValue model_ae50();
ExcelValue model_af50();
ExcelValue model_ag50();
ExcelValue model_ah50();
ExcelValue model_ai50();
ExcelValue model_aj50();
ExcelValue model_ak50();
ExcelValue model_al50();
ExcelValue model_am50();
ExcelValue model_an50();
ExcelValue model_b51();
ExcelValue model_c51();
ExcelValue model_d51();
ExcelValue model_e51();
ExcelValue model_f51();
ExcelValue model_g51();
ExcelValue model_h51();
ExcelValue model_i51();
ExcelValue model_j51();
ExcelValue model_k51();
ExcelValue model_l51();
ExcelValue model_m51();
ExcelValue model_n51();
ExcelValue model_o51();
ExcelValue model_p51();
ExcelValue model_q51();
ExcelValue model_r51();
ExcelValue model_s51();
ExcelValue model_t51();
ExcelValue model_u51();
ExcelValue model_v51();
ExcelValue model_w51();
ExcelValue model_x51();
ExcelValue model_y51();
ExcelValue model_z51();
ExcelValue model_aa51();
ExcelValue model_ab51();
ExcelValue model_ac51();
ExcelValue model_ad51();
ExcelValue model_ae51();
ExcelValue model_af51();
ExcelValue model_ag51();
ExcelValue model_ah51();
ExcelValue model_ai51();
ExcelValue model_aj51();
ExcelValue model_ak51();
ExcelValue model_al51();
ExcelValue model_am51();
ExcelValue model_an51();
ExcelValue model_b52();
ExcelValue model_c52();
ExcelValue model_d52();
ExcelValue model_e52();
ExcelValue model_f52();
ExcelValue model_g52();
ExcelValue model_h52();
ExcelValue model_i52();
ExcelValue model_j52();
ExcelValue model_k52();
ExcelValue model_l52();
ExcelValue model_m52();
ExcelValue model_n52();
ExcelValue model_o52();
ExcelValue model_p52();
ExcelValue model_q52();
ExcelValue model_r52();
ExcelValue model_s52();
ExcelValue model_t52();
ExcelValue model_u52();
ExcelValue model_v52();
ExcelValue model_w52();
ExcelValue model_x52();
ExcelValue model_y52();
ExcelValue model_z52();
ExcelValue model_aa52();
ExcelValue model_ab52();
ExcelValue model_ac52();
ExcelValue model_ad52();
ExcelValue model_ae52();
ExcelValue model_af52();
ExcelValue model_ag52();
ExcelValue model_ah52();
ExcelValue model_ai52();
ExcelValue model_aj52();
ExcelValue model_ak52();
ExcelValue model_al52();
ExcelValue model_am52();
ExcelValue model_an52();
ExcelValue model_b53();
ExcelValue model_c53();
ExcelValue model_d53();
ExcelValue model_e53();
ExcelValue model_f53();
ExcelValue model_g53();
ExcelValue model_h53();
ExcelValue model_i53();
ExcelValue model_j53();
ExcelValue model_k53();
ExcelValue model_l53();
ExcelValue model_m53();
ExcelValue model_n53();
ExcelValue model_o53();
ExcelValue model_p53();
ExcelValue model_q53();
ExcelValue model_r53();
ExcelValue model_s53();
ExcelValue model_t53();
ExcelValue model_u53();
ExcelValue model_v53();
ExcelValue model_w53();
ExcelValue model_x53();
ExcelValue model_y53();
ExcelValue model_z53();
ExcelValue model_aa53();
ExcelValue model_ab53();
ExcelValue model_ac53();
ExcelValue model_ad53();
ExcelValue model_ae53();
ExcelValue model_af53();
ExcelValue model_ag53();
ExcelValue model_ah53();
ExcelValue model_ai53();
ExcelValue model_aj53();
ExcelValue model_ak53();
ExcelValue model_al53();
ExcelValue model_am53();
ExcelValue model_an53();
ExcelValue model_c54();
ExcelValue model_d54();
ExcelValue model_e54();
ExcelValue model_f54();
ExcelValue model_g54();
ExcelValue model_h54();
ExcelValue model_i54();
ExcelValue model_j54();
ExcelValue model_k54();
ExcelValue model_l54();
ExcelValue model_m54();
ExcelValue model_n54();
ExcelValue model_o54();
ExcelValue model_p54();
ExcelValue model_q54();
ExcelValue model_r54();
ExcelValue model_s54();
ExcelValue model_t54();
ExcelValue model_u54();
ExcelValue model_v54();
ExcelValue model_w54();
ExcelValue model_x54();
ExcelValue model_y54();
ExcelValue model_z54();
ExcelValue model_aa54();
ExcelValue model_ab54();
ExcelValue model_ac54();
ExcelValue model_ad54();
ExcelValue model_ae54();
ExcelValue model_af54();
ExcelValue model_ag54();
ExcelValue model_ah54();
ExcelValue model_ai54();
ExcelValue model_aj54();
ExcelValue model_ak54();
ExcelValue model_al54();
ExcelValue model_am54();
ExcelValue model_an54();
ExcelValue model_b55();
ExcelValue model_c55();
ExcelValue model_d55();
ExcelValue model_e55();
ExcelValue model_f55();
ExcelValue model_g55();
ExcelValue model_h55();
ExcelValue model_i55();
ExcelValue model_j55();
ExcelValue model_k55();
ExcelValue model_l55();
ExcelValue model_m55();
ExcelValue model_n55();
ExcelValue model_o55();
ExcelValue model_p55();
ExcelValue model_q55();
ExcelValue model_r55();
ExcelValue model_s55();
ExcelValue model_t55();
ExcelValue model_u55();
ExcelValue model_v55();
ExcelValue model_w55();
ExcelValue model_x55();
ExcelValue model_y55();
ExcelValue model_z55();
ExcelValue model_aa55();
ExcelValue model_ab55();
ExcelValue model_ac55();
ExcelValue model_ad55();
ExcelValue model_ae55();
ExcelValue model_af55();
ExcelValue model_ag55();
ExcelValue model_ah55();
ExcelValue model_ai55();
ExcelValue model_aj55();
ExcelValue model_ak55();
ExcelValue model_al55();
ExcelValue model_am55();
ExcelValue model_an55();
ExcelValue model_c56();
ExcelValue model_d56();
ExcelValue model_e56();
ExcelValue model_f56();
ExcelValue model_g56();
ExcelValue model_h56();
ExcelValue model_i56();
ExcelValue model_j56();
ExcelValue model_k56();
ExcelValue model_l56();
ExcelValue model_m56();
ExcelValue model_n56();
ExcelValue model_o56();
ExcelValue model_p56();
ExcelValue model_q56();
ExcelValue model_r56();
ExcelValue model_s56();
ExcelValue model_t56();
ExcelValue model_u56();
ExcelValue model_v56();
ExcelValue model_w56();
ExcelValue model_x56();
ExcelValue model_y56();
ExcelValue model_z56();
ExcelValue model_aa56();
ExcelValue model_ab56();
ExcelValue model_ac56();
ExcelValue model_ad56();
ExcelValue model_ae56();
ExcelValue model_af56();
ExcelValue model_ag56();
ExcelValue model_ah56();
ExcelValue model_ai56();
ExcelValue model_aj56();
ExcelValue model_ak56();
ExcelValue model_al56();
ExcelValue model_am56();
ExcelValue model_an56();
static ExcelValue model_d59();
static ExcelValue model_e59();
static ExcelValue model_f59();
static ExcelValue model_g59();
static ExcelValue model_h59();
static ExcelValue model_i59();
static ExcelValue model_j59();
static ExcelValue model_k59();
static ExcelValue model_l59();
static ExcelValue model_m59();
static ExcelValue model_n59();
static ExcelValue model_o59();
static ExcelValue model_p59();
static ExcelValue model_q59();
static ExcelValue model_r59();
static ExcelValue model_s59();
static ExcelValue model_t59();
static ExcelValue model_u59();
static ExcelValue model_v59();
static ExcelValue model_w59();
static ExcelValue model_x59();
static ExcelValue model_y59();
static ExcelValue model_z59();
static ExcelValue model_aa59();
static ExcelValue model_ab59();
static ExcelValue model_ac59();
static ExcelValue model_ad59();
static ExcelValue model_ae59();
static ExcelValue model_af59();
static ExcelValue model_ag59();
static ExcelValue model_ah59();
static ExcelValue model_ai59();
static ExcelValue model_aj59();
static ExcelValue model_ak59();
static ExcelValue model_al59();
static ExcelValue model_am59();
static ExcelValue model_d60();
static ExcelValue model_e60();
static ExcelValue model_f60();
static ExcelValue model_g60();
static ExcelValue model_h60();
static ExcelValue model_i60();
static ExcelValue model_j60();
static ExcelValue model_k60();
static ExcelValue model_l60();
static ExcelValue model_m60();
static ExcelValue model_n60();
static ExcelValue model_o60();
static ExcelValue model_p60();
static ExcelValue model_q60();
static ExcelValue model_r60();
static ExcelValue model_s60();
static ExcelValue model_t60();
static ExcelValue model_u60();
static ExcelValue model_v60();
static ExcelValue model_w60();
static ExcelValue model_x60();
static ExcelValue model_y60();
static ExcelValue model_z60();
static ExcelValue model_aa60();
static ExcelValue model_ab60();
static ExcelValue model_ac60();
static ExcelValue model_ad60();
static ExcelValue model_ae60();
static ExcelValue model_af60();
static ExcelValue model_ag60();
static ExcelValue model_ah60();
static ExcelValue model_ai60();
static ExcelValue model_aj60();
static ExcelValue model_ak60();
static ExcelValue model_al60();
static ExcelValue model_am60();
static ExcelValue model_c63();
static ExcelValue model_d63();
static ExcelValue model_e63();
static ExcelValue model_f63();
static ExcelValue model_g63();
static ExcelValue model_h63();
static ExcelValue model_i63();
static ExcelValue model_j63();
static ExcelValue model_k63();
static ExcelValue model_l63();
static ExcelValue model_m63();
static ExcelValue model_n63();
static ExcelValue model_o63();
static ExcelValue model_p63();
static ExcelValue model_q63();
static ExcelValue model_r63();
static ExcelValue model_s63();
static ExcelValue model_t63();
static ExcelValue model_u63();
static ExcelValue model_v63();
static ExcelValue model_w63();
static ExcelValue model_x63();
static ExcelValue model_y63();
static ExcelValue model_z63();
static ExcelValue model_aa63();
static ExcelValue model_ab63();
static ExcelValue model_ac63();
static ExcelValue model_ad63();
static ExcelValue model_ae63();
static ExcelValue model_af63();
static ExcelValue model_ag63();
static ExcelValue model_ah63();
static ExcelValue model_ai63();
static ExcelValue model_aj63();
static ExcelValue model_ak63();
static ExcelValue model_al63();
static ExcelValue model_am63();
static ExcelValue model_an63();
static ExcelValue model_c64();
static ExcelValue model_d64();
static ExcelValue model_e64();
static ExcelValue model_f64();
static ExcelValue model_g64();
static ExcelValue model_h64();
static ExcelValue model_i64();
static ExcelValue model_j64();
static ExcelValue model_k64();
static ExcelValue model_l64();
static ExcelValue model_m64();
static ExcelValue model_n64();
static ExcelValue model_o64();
static ExcelValue model_p64();
static ExcelValue model_q64();
static ExcelValue model_r64();
static ExcelValue model_s64();
static ExcelValue model_t64();
static ExcelValue model_u64();
static ExcelValue model_v64();
static ExcelValue model_w64();
static ExcelValue model_x64();
static ExcelValue model_y64();
static ExcelValue model_z64();
static ExcelValue model_aa64();
static ExcelValue model_ab64();
static ExcelValue model_ac64();
static ExcelValue model_ad64();
static ExcelValue model_ae64();
static ExcelValue model_af64();
static ExcelValue model_ag64();
static ExcelValue model_ah64();
static ExcelValue model_ai64();
static ExcelValue model_aj64();
static ExcelValue model_ak64();
static ExcelValue model_al64();
static ExcelValue model_am64();
static ExcelValue model_an64();
static ExcelValue model_b67();
static ExcelValue model_c67();
static ExcelValue model_d67();
static ExcelValue model_e67();
static ExcelValue model_f67();
static ExcelValue model_g67();
static ExcelValue model_h67();
static ExcelValue model_i67();
static ExcelValue model_j67();
static ExcelValue model_k67();
static ExcelValue model_l67();
static ExcelValue model_m67();
static ExcelValue model_n67();
static ExcelValue model_o67();
static ExcelValue model_p67();
static ExcelValue model_q67();
static ExcelValue model_r67();
static ExcelValue model_s67();
static ExcelValue model_t67();
static ExcelValue model_u67();
static ExcelValue model_v67();
static ExcelValue model_w67();
static ExcelValue model_x67();
static ExcelValue model_y67();
static ExcelValue model_z67();
static ExcelValue model_aa67();
static ExcelValue model_ab67();
static ExcelValue model_ac67();
static ExcelValue model_ad67();
static ExcelValue model_ae67();
static ExcelValue model_af67();
static ExcelValue model_ag67();
static ExcelValue model_ah67();
static ExcelValue model_ai67();
static ExcelValue model_aj67();
static ExcelValue model_ak67();
static ExcelValue model_al67();
static ExcelValue model_am67();
static ExcelValue model_an67();
static ExcelValue model_b68();
static ExcelValue model_c68();
static ExcelValue model_d68();
static ExcelValue model_e68();
static ExcelValue model_f68();
static ExcelValue model_g68();
static ExcelValue model_h68();
static ExcelValue model_i68();
static ExcelValue model_j68();
static ExcelValue model_k68();
static ExcelValue model_l68();
static ExcelValue model_m68();
static ExcelValue model_n68();
static ExcelValue model_o68();
static ExcelValue model_p68();
static ExcelValue model_q68();
static ExcelValue model_r68();
static ExcelValue model_s68();
static ExcelValue model_t68();
static ExcelValue model_u68();
static ExcelValue model_v68();
static ExcelValue model_w68();
static ExcelValue model_x68();
static ExcelValue model_y68();
static ExcelValue model_z68();
static ExcelValue model_aa68();
static ExcelValue model_ab68();
static ExcelValue model_ac68();
static ExcelValue model_ad68();
static ExcelValue model_ae68();
static ExcelValue model_af68();
static ExcelValue model_ag68();
static ExcelValue model_ah68();
static ExcelValue model_ai68();
static ExcelValue model_aj68();
static ExcelValue model_ak68();
static ExcelValue model_al68();
static ExcelValue model_am68();
static ExcelValue model_an68();
static ExcelValue model_b72();
static ExcelValue model_c72();
static ExcelValue model_d72();
static ExcelValue model_e72();
static ExcelValue model_f72();
static ExcelValue model_g72();
static ExcelValue model_h72();
static ExcelValue model_i72();
static ExcelValue model_j72();
static ExcelValue model_k72();
static ExcelValue model_l72();
static ExcelValue model_m72();
static ExcelValue model_n72();
static ExcelValue model_o72();
static ExcelValue model_p72();
static ExcelValue model_q72();
static ExcelValue model_r72();
static ExcelValue model_s72();
static ExcelValue model_t72();
static ExcelValue model_u72();
static ExcelValue model_v72();
static ExcelValue model_w72();
static ExcelValue model_x72();
static ExcelValue model_y72();
static ExcelValue model_z72();
static ExcelValue model_aa72();
static ExcelValue model_ab72();
static ExcelValue model_ac72();
static ExcelValue model_ad72();
static ExcelValue model_ae72();
static ExcelValue model_af72();
static ExcelValue model_ag72();
static ExcelValue model_ah72();
static ExcelValue model_ai72();
static ExcelValue model_aj72();
static ExcelValue model_ak72();
static ExcelValue model_al72();
static ExcelValue model_am72();
static ExcelValue model_an72();
static ExcelValue model_k74();
static ExcelValue model_l74();
static ExcelValue model_m74();
static ExcelValue model_n74();
static ExcelValue model_o74();
static ExcelValue model_p74();
static ExcelValue model_q74();
static ExcelValue model_r74();
static ExcelValue model_s74();
static ExcelValue model_t74();
static ExcelValue model_u74();
static ExcelValue model_v74();
static ExcelValue model_w74();
static ExcelValue model_x74();
static ExcelValue model_y74();
static ExcelValue model_z74();
static ExcelValue model_aa74();
static ExcelValue model_ab74();
static ExcelValue model_ac74();
static ExcelValue model_ad74();
static ExcelValue model_ae74();
static ExcelValue model_af74();
static ExcelValue model_ag74();
static ExcelValue model_ah74();
static ExcelValue model_ai74();
static ExcelValue model_aj74();
static ExcelValue model_ak74();
static ExcelValue model_al74();
static ExcelValue model_am74();
static ExcelValue model_an74();
static ExcelValue model_b75();
static ExcelValue model_c75();
static ExcelValue model_d75();
static ExcelValue model_e75();
static ExcelValue model_f75();
static ExcelValue model_g75();
static ExcelValue model_h75();
static ExcelValue model_i75();
static ExcelValue model_j75();
static ExcelValue model_k75();
static ExcelValue model_l75();
static ExcelValue model_m75();
static ExcelValue model_n75();
static ExcelValue model_o75();
static ExcelValue model_p75();
static ExcelValue model_q75();
static ExcelValue model_r75();
static ExcelValue model_s75();
static ExcelValue model_t75();
static ExcelValue model_u75();
static ExcelValue model_v75();
static ExcelValue model_w75();
static ExcelValue model_x75();
static ExcelValue model_y75();
static ExcelValue model_z75();
static ExcelValue model_aa75();
static ExcelValue model_ab75();
static ExcelValue model_ac75();
static ExcelValue model_ad75();
static ExcelValue model_ae75();
static ExcelValue model_af75();
static ExcelValue model_ag75();
static ExcelValue model_ah75();
static ExcelValue model_ai75();
static ExcelValue model_aj75();
static ExcelValue model_ak75();
static ExcelValue model_al75();
static ExcelValue model_am75();
static ExcelValue model_an75();
static ExcelValue model_b76();
static ExcelValue model_c76();
static ExcelValue model_d76();
static ExcelValue model_e76();
static ExcelValue model_f76();
static ExcelValue model_g76();
static ExcelValue model_h76();
static ExcelValue model_i76();
static ExcelValue model_j76();
static ExcelValue model_k76();
static ExcelValue model_l76();
static ExcelValue model_m76();
static ExcelValue model_n76();
static ExcelValue model_o76();
static ExcelValue model_p76();
static ExcelValue model_q76();
static ExcelValue model_r76();
static ExcelValue model_s76();
static ExcelValue model_t76();
static ExcelValue model_u76();
static ExcelValue model_v76();
static ExcelValue model_w76();
static ExcelValue model_x76();
static ExcelValue model_y76();
static ExcelValue model_z76();
static ExcelValue model_aa76();
static ExcelValue model_ab76();
static ExcelValue model_ac76();
static ExcelValue model_ad76();
static ExcelValue model_ae76();
static ExcelValue model_af76();
static ExcelValue model_ag76();
static ExcelValue model_ah76();
static ExcelValue model_ai76();
static ExcelValue model_aj76();
static ExcelValue model_ak76();
static ExcelValue model_al76();
static ExcelValue model_am76();
static ExcelValue model_an76();
static ExcelValue model_b77();
static ExcelValue model_c77();
static ExcelValue model_d77();
static ExcelValue model_e77();
static ExcelValue model_f77();
static ExcelValue model_g77();
static ExcelValue model_h77();
static ExcelValue model_i77();
static ExcelValue model_j77();
static ExcelValue model_k77();
static ExcelValue model_l77();
static ExcelValue model_m77();
static ExcelValue model_n77();
static ExcelValue model_o77();
static ExcelValue model_p77();
static ExcelValue model_q77();
static ExcelValue model_r77();
static ExcelValue model_s77();
static ExcelValue model_t77();
static ExcelValue model_u77();
static ExcelValue model_v77();
static ExcelValue model_w77();
static ExcelValue model_x77();
static ExcelValue model_y77();
static ExcelValue model_z77();
static ExcelValue model_aa77();
static ExcelValue model_ab77();
static ExcelValue model_ac77();
static ExcelValue model_ad77();
static ExcelValue model_ae77();
static ExcelValue model_af77();
static ExcelValue model_ag77();
static ExcelValue model_ah77();
static ExcelValue model_ai77();
static ExcelValue model_aj77();
static ExcelValue model_ak77();
static ExcelValue model_al77();
static ExcelValue model_am77();
static ExcelValue model_an77();
ExcelValue model_b85();
ExcelValue model_c85();
ExcelValue model_d85();
ExcelValue model_e85();
ExcelValue model_f85();
ExcelValue model_g85();
ExcelValue model_h85();
ExcelValue model_i85();
ExcelValue model_j85();
ExcelValue model_k85();
ExcelValue model_l85();
ExcelValue model_m85();
ExcelValue model_n85();
ExcelValue model_o85();
ExcelValue model_p85();
ExcelValue model_q85();
ExcelValue model_r85();
ExcelValue model_s85();
ExcelValue model_t85();
ExcelValue model_u85();
ExcelValue model_v85();
ExcelValue model_w85();
ExcelValue model_x85();
ExcelValue model_y85();
ExcelValue model_z85();
ExcelValue model_aa85();
ExcelValue model_ab85();
ExcelValue model_ac85();
ExcelValue model_ad85();
ExcelValue model_ae85();
ExcelValue model_af85();
ExcelValue model_ag85();
ExcelValue model_ah85();
ExcelValue model_ai85();
ExcelValue model_aj85();
ExcelValue model_ak85();
ExcelValue model_al85();
ExcelValue model_am85();
ExcelValue model_an85();
static ExcelValue model_k86();
static ExcelValue model_l86();
static ExcelValue model_m86();
static ExcelValue model_n86();
static ExcelValue model_o86();
static ExcelValue model_p86();
static ExcelValue model_q86();
static ExcelValue model_r86();
static ExcelValue model_s86();
static ExcelValue model_t86();
static ExcelValue model_u86();
static ExcelValue model_v86();
static ExcelValue model_w86();
static ExcelValue model_x86();
static ExcelValue model_y86();
static ExcelValue model_z86();
static ExcelValue model_aa86();
static ExcelValue model_ab86();
static ExcelValue model_ac86();
static ExcelValue model_ad86();
static ExcelValue model_ae86();
static ExcelValue model_af86();
static ExcelValue model_ag86();
static ExcelValue model_ah86();
static ExcelValue model_ai86();
static ExcelValue model_aj86();
static ExcelValue model_ak86();
static ExcelValue model_al86();
static ExcelValue model_am86();
static ExcelValue model_an86();
ExcelValue model_b89();
ExcelValue model_c89();
ExcelValue model_d89();
ExcelValue model_e89();
ExcelValue model_f89();
ExcelValue model_g89();
ExcelValue model_h89();
ExcelValue model_i89();
ExcelValue model_j89();
ExcelValue model_k89();
ExcelValue model_l89();
ExcelValue model_m89();
ExcelValue model_n89();
ExcelValue model_o89();
ExcelValue model_p89();
ExcelValue model_q89();
ExcelValue model_r89();
ExcelValue model_s89();
ExcelValue model_t89();
ExcelValue model_u89();
ExcelValue model_v89();
ExcelValue model_w89();
ExcelValue model_x89();
ExcelValue model_y89();
ExcelValue model_z89();
ExcelValue model_aa89();
ExcelValue model_ab89();
ExcelValue model_ac89();
ExcelValue model_ad89();
ExcelValue model_ae89();
ExcelValue model_af89();
ExcelValue model_ag89();
ExcelValue model_ah89();
ExcelValue model_ai89();
ExcelValue model_aj89();
ExcelValue model_ak89();
ExcelValue model_al89();
ExcelValue model_am89();
ExcelValue model_an89();
ExcelValue model_b56();
ExcelValue model_b54();
// end of definitions

// Used to decide whether to recalculate a cell
static int variable_set[1404];

// Used to reset all cached values and free up memory
void reset() {
  int i;
  cell_counter = 0;
  free_all_allocated_memory(); 
  for(i = 0; i < 1404; i++) {
    variable_set[i] = 0;
  }
};

// starting the value constants
static ExcelValue C1 = {.type = ExcelNumber, .number = 2020};
static ExcelValue C2 = {.type = ExcelNumber, .number = 613};
static ExcelValue C3 = {.type = ExcelNumber, .number = 0.3};
static ExcelValue C4 = {.type = ExcelNumber, .number = 2021};
static ExcelValue C5 = {.type = ExcelNumber, .number = 43.83};
static ExcelValue C6 = {.type = ExcelNumber, .number = 4.383};
static ExcelValue C7 = {.type = ExcelNumber, .number = 0.5};
static ExcelValue C8 = {.type = ExcelNumber, .number = 1};
static ExcelValue C9 = {.type = ExcelNumber, .number = 30};
static ExcelValue C10 = {.type = ExcelNumber, .number = 346};
static ExcelValue C11 = {.type = ExcelNumber, .number = -0.004402503205486741};
static ExcelValue C12 = {.type = ExcelNumber, .number = 38.7};
static ExcelValue C13 = {.type = ExcelNumber, .number = 59};
static ExcelValue C14 = {.type = ExcelNumber, .number = 10};
static ExcelValue C15 = {.type = ExcelNumber, .number = 5.4};
static ExcelValue C16 = {.type = ExcelNumber, .number = 69};
static ExcelValue C17 = {.type = ExcelNumber, .number = 97.7};
static ExcelValue C18 = {.type = ExcelNumber, .number = 8};
static ExcelValue C19 = {.type = ExcelNumber, .number = 650};
static ExcelValue C20 = {.type = ExcelNumber, .number = 370};
static ExcelValue C21 = {.type = ExcelNumber, .number = 350};
static ExcelValue C22 = {.type = ExcelNumber, .number = 1.5};
static ExcelValue C23 = {.type = ExcelNumber, .number = 2};
static ExcelValue C24 = {.type = ExcelNumber, .number = 2015};
static ExcelValue C25 = {.type = ExcelNumber, .number = 344.4767338909016};
static ExcelValue C26 = {.type = ExcelNumber, .number = 0.9955974967945133};
static ExcelValue C27 = {.type = ExcelNumber, .number = 1.0204470321495855};
static ExcelValue C28 = {.type = ExcelNumber, .number = 248.3};
static ExcelValue C29 = {.type = ExcelNumber, .number = 615.0};
static ExcelValue C30 = {.type = ExcelNumber, .number = 580.0};
static ExcelValue C31 = {.type = ExcelNumber, .number = -35.0};
static ExcelValue C32 = {.type = ExcelNumber, .number = -0.6666666666666666};
static ExcelValue C33 = {.type = ExcelNumber, .number = 161.395};
static ExcelValue C34 = {.type = ExcelNumber, .number = 1000};
static ExcelValue C35 = {.type = ExcelNumber, .number = 3.256666666666667};
static ExcelValue C36 = {.type = ExcelNumber, .number = 0.03333333333333333};
static ExcelValue C37 = {.type = ExcelNumber, .number = 0};
static ExcelValue C38 = {.type = ExcelNumber, .number = 1.0};
static ExcelValue C39 = {.type = ExcelNumber, .number = 1.0096355280556115};
static ExcelValue C40 = {.type = ExcelNumber, .number = 519.0};
static ExcelValue C41 = {.type = ExcelNumber, .number = 173.0};
// ending the value constants

// starting common elements
static ExcelValue _common0() {
  static ExcelValue result;
  if(variable_set[0] == 1) { return result;}
  result = multiply(divide(divide(multiply(model_t51(),subtract(model_t48(),model_t86())),C34),model_t48()),C34);
  variable_set[0] = 1;
  return result;
}

static ExcelValue _common1() {
  static ExcelValue result;
  if(variable_set[1] == 1) { return result;}
  result = divide(divide(multiply(model_t51(),subtract(model_t48(),model_t86())),C34),model_t48());
  variable_set[1] = 1;
  return result;
}

static ExcelValue _common2() {
  static ExcelValue result;
  if(variable_set[2] == 1) { return result;}
  result = divide(multiply(model_t51(),subtract(model_t48(),model_t86())),C34);
  variable_set[2] = 1;
  return result;
}

static ExcelValue _common3() {
  static ExcelValue result;
  if(variable_set[3] == 1) { return result;}
  result = multiply(model_t51(),subtract(model_t48(),model_t86()));
  variable_set[3] = 1;
  return result;
}

static ExcelValue _common4() {
  static ExcelValue result;
  if(variable_set[4] == 1) { return result;}
  result = subtract(model_t48(),model_t86());
  variable_set[4] = 1;
  return result;
}

static ExcelValue _common5() {
  static ExcelValue result;
  if(variable_set[5] == 1) { return result;}
  result = multiply(divide(divide(multiply(add(model_am51(),C32),subtract(model_an48(),model_an86())),C34),model_an48()),C34);
  variable_set[5] = 1;
  return result;
}

static ExcelValue _common6() {
  static ExcelValue result;
  if(variable_set[6] == 1) { return result;}
  result = divide(divide(multiply(add(model_am51(),C32),subtract(model_an48(),model_an86())),C34),model_an48());
  variable_set[6] = 1;
  return result;
}

static ExcelValue _common7() {
  static ExcelValue result;
  if(variable_set[7] == 1) { return result;}
  result = divide(multiply(add(model_am51(),C32),subtract(model_an48(),model_an86())),C34);
  variable_set[7] = 1;
  return result;
}

static ExcelValue _common8() {
  static ExcelValue result;
  if(variable_set[8] == 1) { return result;}
  result = multiply(add(model_am51(),C32),subtract(model_an48(),model_an86()));
  variable_set[8] = 1;
  return result;
}

static ExcelValue _common9() {
  static ExcelValue result;
  if(variable_set[9] == 1) { return result;}
  result = add(model_am51(),C32);
  variable_set[9] = 1;
  return result;
}

static ExcelValue _common10() {
  static ExcelValue result;
  if(variable_set[10] == 1) { return result;}
  result = subtract(model_an48(),model_an86());
  variable_set[10] = 1;
  return result;
}

static ExcelValue _common11() {
  static ExcelValue result;
  if(variable_set[11] == 1) { return result;}
  result = subtract(C25,model_c49());
  variable_set[11] = 1;
  return result;
}

static ExcelValue _common12() {
  static ExcelValue result;
  if(variable_set[12] == 1) { return result;}
  result = subtract(model_d48(),model_d49());
  variable_set[12] = 1;
  return result;
}

static ExcelValue _common13() {
  static ExcelValue result;
  if(variable_set[13] == 1) { return result;}
  result = subtract(model_e48(),model_e49());
  variable_set[13] = 1;
  return result;
}

static ExcelValue _common14() {
  static ExcelValue result;
  if(variable_set[14] == 1) { return result;}
  result = subtract(model_f48(),model_f49());
  variable_set[14] = 1;
  return result;
}

static ExcelValue _common15() {
  static ExcelValue result;
  if(variable_set[15] == 1) { return result;}
  result = subtract(model_g48(),model_g49());
  variable_set[15] = 1;
  return result;
}

static ExcelValue _common16() {
  static ExcelValue result;
  if(variable_set[16] == 1) { return result;}
  result = subtract(model_h48(),model_h49());
  variable_set[16] = 1;
  return result;
}

static ExcelValue _common17() {
  static ExcelValue result;
  if(variable_set[17] == 1) { return result;}
  result = subtract(model_i48(),model_i49());
  variable_set[17] = 1;
  return result;
}

static ExcelValue _common18() {
  static ExcelValue result;
  if(variable_set[18] == 1) { return result;}
  result = subtract(model_j48(),model_j49());
  variable_set[18] = 1;
  return result;
}

static ExcelValue _common19() {
  static ExcelValue result;
  if(variable_set[19] == 1) { return result;}
  result = subtract(model_k48(),model_k86());
  variable_set[19] = 1;
  return result;
}

static ExcelValue _common20() {
  static ExcelValue result;
  if(variable_set[20] == 1) { return result;}
  result = subtract(model_l48(),model_l86());
  variable_set[20] = 1;
  return result;
}

static ExcelValue _common21() {
  static ExcelValue result;
  if(variable_set[21] == 1) { return result;}
  result = subtract(model_m48(),model_m86());
  variable_set[21] = 1;
  return result;
}

static ExcelValue _common22() {
  static ExcelValue result;
  if(variable_set[22] == 1) { return result;}
  result = subtract(model_n48(),model_n86());
  variable_set[22] = 1;
  return result;
}

static ExcelValue _common23() {
  static ExcelValue result;
  if(variable_set[23] == 1) { return result;}
  result = subtract(model_o48(),model_o86());
  variable_set[23] = 1;
  return result;
}

static ExcelValue _common24() {
  static ExcelValue result;
  if(variable_set[24] == 1) { return result;}
  result = subtract(model_p48(),model_p86());
  variable_set[24] = 1;
  return result;
}

static ExcelValue _common25() {
  static ExcelValue result;
  if(variable_set[25] == 1) { return result;}
  result = subtract(model_q48(),model_q86());
  variable_set[25] = 1;
  return result;
}

static ExcelValue _common26() {
  static ExcelValue result;
  if(variable_set[26] == 1) { return result;}
  result = subtract(model_r48(),model_r86());
  variable_set[26] = 1;
  return result;
}

static ExcelValue _common27() {
  static ExcelValue result;
  if(variable_set[27] == 1) { return result;}
  result = subtract(model_s48(),model_s86());
  variable_set[27] = 1;
  return result;
}

static ExcelValue _common28() {
  static ExcelValue result;
  if(variable_set[28] == 1) { return result;}
  result = subtract(model_u48(),model_u86());
  variable_set[28] = 1;
  return result;
}

static ExcelValue _common29() {
  static ExcelValue result;
  if(variable_set[29] == 1) { return result;}
  result = subtract(model_v48(),model_v86());
  variable_set[29] = 1;
  return result;
}

static ExcelValue _common30() {
  static ExcelValue result;
  if(variable_set[30] == 1) { return result;}
  result = subtract(model_w48(),model_w86());
  variable_set[30] = 1;
  return result;
}

static ExcelValue _common31() {
  static ExcelValue result;
  if(variable_set[31] == 1) { return result;}
  result = subtract(model_x48(),model_x86());
  variable_set[31] = 1;
  return result;
}

static ExcelValue _common32() {
  static ExcelValue result;
  if(variable_set[32] == 1) { return result;}
  result = subtract(model_y48(),model_y86());
  variable_set[32] = 1;
  return result;
}

static ExcelValue _common33() {
  static ExcelValue result;
  if(variable_set[33] == 1) { return result;}
  result = subtract(model_z48(),model_z86());
  variable_set[33] = 1;
  return result;
}

static ExcelValue _common34() {
  static ExcelValue result;
  if(variable_set[34] == 1) { return result;}
  result = subtract(model_aa48(),model_aa86());
  variable_set[34] = 1;
  return result;
}

static ExcelValue _common35() {
  static ExcelValue result;
  if(variable_set[35] == 1) { return result;}
  result = subtract(model_ab48(),model_ab86());
  variable_set[35] = 1;
  return result;
}

static ExcelValue _common36() {
  static ExcelValue result;
  if(variable_set[36] == 1) { return result;}
  result = subtract(model_ac48(),model_ac86());
  variable_set[36] = 1;
  return result;
}

static ExcelValue _common37() {
  static ExcelValue result;
  if(variable_set[37] == 1) { return result;}
  result = subtract(model_ad48(),model_ad86());
  variable_set[37] = 1;
  return result;
}

static ExcelValue _common38() {
  static ExcelValue result;
  if(variable_set[38] == 1) { return result;}
  result = subtract(model_ae48(),model_ae86());
  variable_set[38] = 1;
  return result;
}

static ExcelValue _common39() {
  static ExcelValue result;
  if(variable_set[39] == 1) { return result;}
  result = subtract(model_af48(),model_af86());
  variable_set[39] = 1;
  return result;
}

static ExcelValue _common40() {
  static ExcelValue result;
  if(variable_set[40] == 1) { return result;}
  result = subtract(model_ag48(),model_ag86());
  variable_set[40] = 1;
  return result;
}

static ExcelValue _common41() {
  static ExcelValue result;
  if(variable_set[41] == 1) { return result;}
  result = subtract(model_ah48(),model_ah86());
  variable_set[41] = 1;
  return result;
}

static ExcelValue _common42() {
  static ExcelValue result;
  if(variable_set[42] == 1) { return result;}
  result = subtract(model_ai48(),model_ai86());
  variable_set[42] = 1;
  return result;
}

static ExcelValue _common43() {
  static ExcelValue result;
  if(variable_set[43] == 1) { return result;}
  result = subtract(model_aj48(),model_aj86());
  variable_set[43] = 1;
  return result;
}

static ExcelValue _common44() {
  static ExcelValue result;
  if(variable_set[44] == 1) { return result;}
  result = subtract(model_ak48(),model_ak86());
  variable_set[44] = 1;
  return result;
}

static ExcelValue _common45() {
  static ExcelValue result;
  if(variable_set[45] == 1) { return result;}
  result = subtract(model_al48(),model_al86());
  variable_set[45] = 1;
  return result;
}

static ExcelValue _common46() {
  static ExcelValue result;
  if(variable_set[46] == 1) { return result;}
  result = subtract(model_am48(),model_am86());
  variable_set[46] = 1;
  return result;
}

static ExcelValue _common47() {
  static ExcelValue result;
  if(variable_set[47] == 1) { return result;}
  result = divide(multiply(C29,subtract(C25,model_c49())),C34);
  variable_set[47] = 1;
  return result;
}

static ExcelValue _common48() {
  static ExcelValue result;
  if(variable_set[48] == 1) { return result;}
  result = multiply(C29,subtract(C25,model_c49()));
  variable_set[48] = 1;
  return result;
}

static ExcelValue _common49() {
  static ExcelValue result;
  if(variable_set[49] == 1) { return result;}
  result = divide(multiply(C30,subtract(model_d48(),model_d49())),C34);
  variable_set[49] = 1;
  return result;
}

static ExcelValue _common50() {
  static ExcelValue result;
  if(variable_set[50] == 1) { return result;}
  result = multiply(C30,subtract(model_d48(),model_d49()));
  variable_set[50] = 1;
  return result;
}

static ExcelValue _common51() {
  static ExcelValue result;
  if(variable_set[51] == 1) { return result;}
  result = divide(multiply(model_e51(),subtract(model_e48(),model_e49())),C34);
  variable_set[51] = 1;
  return result;
}

static ExcelValue _common52() {
  static ExcelValue result;
  if(variable_set[52] == 1) { return result;}
  result = multiply(model_e51(),subtract(model_e48(),model_e49()));
  variable_set[52] = 1;
  return result;
}

static ExcelValue _common53() {
  static ExcelValue result;
  if(variable_set[53] == 1) { return result;}
  result = divide(multiply(model_f51(),subtract(model_f48(),model_f49())),C34);
  variable_set[53] = 1;
  return result;
}

static ExcelValue _common54() {
  static ExcelValue result;
  if(variable_set[54] == 1) { return result;}
  result = multiply(model_f51(),subtract(model_f48(),model_f49()));
  variable_set[54] = 1;
  return result;
}

static ExcelValue _common55() {
  static ExcelValue result;
  if(variable_set[55] == 1) { return result;}
  result = divide(multiply(model_g51(),subtract(model_g48(),model_g49())),C34);
  variable_set[55] = 1;
  return result;
}

static ExcelValue _common56() {
  static ExcelValue result;
  if(variable_set[56] == 1) { return result;}
  result = multiply(model_g51(),subtract(model_g48(),model_g49()));
  variable_set[56] = 1;
  return result;
}

static ExcelValue _common57() {
  static ExcelValue result;
  if(variable_set[57] == 1) { return result;}
  result = divide(multiply(model_h51(),subtract(model_h48(),model_h49())),C34);
  variable_set[57] = 1;
  return result;
}

static ExcelValue _common58() {
  static ExcelValue result;
  if(variable_set[58] == 1) { return result;}
  result = multiply(model_h51(),subtract(model_h48(),model_h49()));
  variable_set[58] = 1;
  return result;
}

static ExcelValue _common59() {
  static ExcelValue result;
  if(variable_set[59] == 1) { return result;}
  result = divide(multiply(model_i51(),subtract(model_i48(),model_i49())),C34);
  variable_set[59] = 1;
  return result;
}

static ExcelValue _common60() {
  static ExcelValue result;
  if(variable_set[60] == 1) { return result;}
  result = multiply(model_i51(),subtract(model_i48(),model_i49()));
  variable_set[60] = 1;
  return result;
}

static ExcelValue _common61() {
  static ExcelValue result;
  if(variable_set[61] == 1) { return result;}
  result = divide(multiply(model_j51(),subtract(model_j48(),model_j49())),C34);
  variable_set[61] = 1;
  return result;
}

static ExcelValue _common62() {
  static ExcelValue result;
  if(variable_set[62] == 1) { return result;}
  result = multiply(model_j51(),subtract(model_j48(),model_j49()));
  variable_set[62] = 1;
  return result;
}

static ExcelValue _common63() {
  static ExcelValue result;
  if(variable_set[63] == 1) { return result;}
  result = divide(multiply(model_k51(),subtract(model_k48(),model_k86())),C34);
  variable_set[63] = 1;
  return result;
}

static ExcelValue _common64() {
  static ExcelValue result;
  if(variable_set[64] == 1) { return result;}
  result = multiply(model_k51(),subtract(model_k48(),model_k86()));
  variable_set[64] = 1;
  return result;
}

static ExcelValue _common65() {
  static ExcelValue result;
  if(variable_set[65] == 1) { return result;}
  result = divide(multiply(model_l51(),subtract(model_l48(),model_l86())),C34);
  variable_set[65] = 1;
  return result;
}

static ExcelValue _common66() {
  static ExcelValue result;
  if(variable_set[66] == 1) { return result;}
  result = multiply(model_l51(),subtract(model_l48(),model_l86()));
  variable_set[66] = 1;
  return result;
}

static ExcelValue _common67() {
  static ExcelValue result;
  if(variable_set[67] == 1) { return result;}
  result = divide(multiply(model_r51(),subtract(model_r48(),model_r86())),C34);
  variable_set[67] = 1;
  return result;
}

static ExcelValue _common68() {
  static ExcelValue result;
  if(variable_set[68] == 1) { return result;}
  result = multiply(model_r51(),subtract(model_r48(),model_r86()));
  variable_set[68] = 1;
  return result;
}

static ExcelValue _common69() {
  static ExcelValue result;
  if(variable_set[69] == 1) { return result;}
  result = divide(multiply(model_s51(),subtract(model_s48(),model_s86())),C34);
  variable_set[69] = 1;
  return result;
}

static ExcelValue _common70() {
  static ExcelValue result;
  if(variable_set[70] == 1) { return result;}
  result = multiply(model_s51(),subtract(model_s48(),model_s86()));
  variable_set[70] = 1;
  return result;
}

static ExcelValue _common71() {
  static ExcelValue result;
  if(variable_set[71] == 1) { return result;}
  result = divide(multiply(model_u51(),subtract(model_u48(),model_u86())),C34);
  variable_set[71] = 1;
  return result;
}

static ExcelValue _common72() {
  static ExcelValue result;
  if(variable_set[72] == 1) { return result;}
  result = multiply(model_u51(),subtract(model_u48(),model_u86()));
  variable_set[72] = 1;
  return result;
}

static ExcelValue _common73() {
  static ExcelValue result;
  if(variable_set[73] == 1) { return result;}
  result = divide(multiply(model_v51(),subtract(model_v48(),model_v86())),C34);
  variable_set[73] = 1;
  return result;
}

static ExcelValue _common74() {
  static ExcelValue result;
  if(variable_set[74] == 1) { return result;}
  result = multiply(model_v51(),subtract(model_v48(),model_v86()));
  variable_set[74] = 1;
  return result;
}

static ExcelValue _common75() {
  static ExcelValue result;
  if(variable_set[75] == 1) { return result;}
  result = divide(multiply(model_w51(),subtract(model_w48(),model_w86())),C34);
  variable_set[75] = 1;
  return result;
}

static ExcelValue _common76() {
  static ExcelValue result;
  if(variable_set[76] == 1) { return result;}
  result = multiply(model_w51(),subtract(model_w48(),model_w86()));
  variable_set[76] = 1;
  return result;
}

static ExcelValue _common77() {
  static ExcelValue result;
  if(variable_set[77] == 1) { return result;}
  result = divide(multiply(model_x51(),subtract(model_x48(),model_x86())),C34);
  variable_set[77] = 1;
  return result;
}

static ExcelValue _common78() {
  static ExcelValue result;
  if(variable_set[78] == 1) { return result;}
  result = multiply(model_x51(),subtract(model_x48(),model_x86()));
  variable_set[78] = 1;
  return result;
}

static ExcelValue _common79() {
  static ExcelValue result;
  if(variable_set[79] == 1) { return result;}
  result = divide(multiply(model_y51(),subtract(model_y48(),model_y86())),C34);
  variable_set[79] = 1;
  return result;
}

static ExcelValue _common80() {
  static ExcelValue result;
  if(variable_set[80] == 1) { return result;}
  result = multiply(model_y51(),subtract(model_y48(),model_y86()));
  variable_set[80] = 1;
  return result;
}

static ExcelValue _common81() {
  static ExcelValue result;
  if(variable_set[81] == 1) { return result;}
  result = divide(multiply(model_z51(),subtract(model_z48(),model_z86())),C34);
  variable_set[81] = 1;
  return result;
}

static ExcelValue _common82() {
  static ExcelValue result;
  if(variable_set[82] == 1) { return result;}
  result = multiply(model_z51(),subtract(model_z48(),model_z86()));
  variable_set[82] = 1;
  return result;
}

static ExcelValue _common83() {
  static ExcelValue result;
  if(variable_set[83] == 1) { return result;}
  result = divide(multiply(model_aa51(),subtract(model_aa48(),model_aa86())),C34);
  variable_set[83] = 1;
  return result;
}

static ExcelValue _common84() {
  static ExcelValue result;
  if(variable_set[84] == 1) { return result;}
  result = multiply(model_aa51(),subtract(model_aa48(),model_aa86()));
  variable_set[84] = 1;
  return result;
}

static ExcelValue _common85() {
  static ExcelValue result;
  if(variable_set[85] == 1) { return result;}
  result = divide(multiply(model_ab51(),subtract(model_ab48(),model_ab86())),C34);
  variable_set[85] = 1;
  return result;
}

static ExcelValue _common86() {
  static ExcelValue result;
  if(variable_set[86] == 1) { return result;}
  result = multiply(model_ab51(),subtract(model_ab48(),model_ab86()));
  variable_set[86] = 1;
  return result;
}

static ExcelValue _common87() {
  static ExcelValue result;
  if(variable_set[87] == 1) { return result;}
  result = divide(multiply(model_ac51(),subtract(model_ac48(),model_ac86())),C34);
  variable_set[87] = 1;
  return result;
}

static ExcelValue _common88() {
  static ExcelValue result;
  if(variable_set[88] == 1) { return result;}
  result = multiply(model_ac51(),subtract(model_ac48(),model_ac86()));
  variable_set[88] = 1;
  return result;
}

static ExcelValue _common89() {
  static ExcelValue result;
  if(variable_set[89] == 1) { return result;}
  result = divide(multiply(model_ad51(),subtract(model_ad48(),model_ad86())),C34);
  variable_set[89] = 1;
  return result;
}

static ExcelValue _common90() {
  static ExcelValue result;
  if(variable_set[90] == 1) { return result;}
  result = multiply(model_ad51(),subtract(model_ad48(),model_ad86()));
  variable_set[90] = 1;
  return result;
}

static ExcelValue _common91() {
  static ExcelValue result;
  if(variable_set[91] == 1) { return result;}
  result = divide(multiply(model_ae51(),subtract(model_ae48(),model_ae86())),C34);
  variable_set[91] = 1;
  return result;
}

static ExcelValue _common92() {
  static ExcelValue result;
  if(variable_set[92] == 1) { return result;}
  result = multiply(model_ae51(),subtract(model_ae48(),model_ae86()));
  variable_set[92] = 1;
  return result;
}

static ExcelValue _common93() {
  static ExcelValue result;
  if(variable_set[93] == 1) { return result;}
  result = divide(multiply(model_af51(),subtract(model_af48(),model_af86())),C34);
  variable_set[93] = 1;
  return result;
}

static ExcelValue _common94() {
  static ExcelValue result;
  if(variable_set[94] == 1) { return result;}
  result = multiply(model_af51(),subtract(model_af48(),model_af86()));
  variable_set[94] = 1;
  return result;
}

static ExcelValue _common95() {
  static ExcelValue result;
  if(variable_set[95] == 1) { return result;}
  result = divide(multiply(model_ag51(),subtract(model_ag48(),model_ag86())),C34);
  variable_set[95] = 1;
  return result;
}

static ExcelValue _common96() {
  static ExcelValue result;
  if(variable_set[96] == 1) { return result;}
  result = multiply(model_ag51(),subtract(model_ag48(),model_ag86()));
  variable_set[96] = 1;
  return result;
}

static ExcelValue _common97() {
  static ExcelValue result;
  if(variable_set[97] == 1) { return result;}
  result = divide(multiply(model_ah51(),subtract(model_ah48(),model_ah86())),C34);
  variable_set[97] = 1;
  return result;
}

static ExcelValue _common98() {
  static ExcelValue result;
  if(variable_set[98] == 1) { return result;}
  result = multiply(model_ah51(),subtract(model_ah48(),model_ah86()));
  variable_set[98] = 1;
  return result;
}

static ExcelValue _common99() {
  static ExcelValue result;
  if(variable_set[99] == 1) { return result;}
  result = divide(multiply(model_ai51(),subtract(model_ai48(),model_ai86())),C34);
  variable_set[99] = 1;
  return result;
}

static ExcelValue _common100() {
  static ExcelValue result;
  if(variable_set[100] == 1) { return result;}
  result = multiply(model_ai51(),subtract(model_ai48(),model_ai86()));
  variable_set[100] = 1;
  return result;
}

static ExcelValue _common101() {
  static ExcelValue result;
  if(variable_set[101] == 1) { return result;}
  result = divide(multiply(model_aj51(),subtract(model_aj48(),model_aj86())),C34);
  variable_set[101] = 1;
  return result;
}

static ExcelValue _common102() {
  static ExcelValue result;
  if(variable_set[102] == 1) { return result;}
  result = multiply(model_aj51(),subtract(model_aj48(),model_aj86()));
  variable_set[102] = 1;
  return result;
}

static ExcelValue _common103() {
  static ExcelValue result;
  if(variable_set[103] == 1) { return result;}
  result = divide(multiply(model_ak51(),subtract(model_ak48(),model_ak86())),C34);
  variable_set[103] = 1;
  return result;
}

static ExcelValue _common104() {
  static ExcelValue result;
  if(variable_set[104] == 1) { return result;}
  result = multiply(model_ak51(),subtract(model_ak48(),model_ak86()));
  variable_set[104] = 1;
  return result;
}

static ExcelValue _common105() {
  static ExcelValue result;
  if(variable_set[105] == 1) { return result;}
  result = divide(multiply(model_al51(),subtract(model_al48(),model_al86())),C34);
  variable_set[105] = 1;
  return result;
}

static ExcelValue _common106() {
  static ExcelValue result;
  if(variable_set[106] == 1) { return result;}
  result = multiply(model_al51(),subtract(model_al48(),model_al86()));
  variable_set[106] = 1;
  return result;
}

static ExcelValue _common107() {
  static ExcelValue result;
  if(variable_set[107] == 1) { return result;}
  result = divide(multiply(model_am51(),subtract(model_am48(),model_am86())),C34);
  variable_set[107] = 1;
  return result;
}

static ExcelValue _common108() {
  static ExcelValue result;
  if(variable_set[108] == 1) { return result;}
  result = multiply(model_am51(),subtract(model_am48(),model_am86()));
  variable_set[108] = 1;
  return result;
}

static ExcelValue _common109() {
  static ExcelValue result;
  if(variable_set[109] == 1) { return result;}
  result = multiply(model_c49(),C36);
  variable_set[109] = 1;
  return result;
}

static ExcelValue _common110() {
  static ExcelValue result;
  if(variable_set[110] == 1) { return result;}
  result = multiply(model_d49(),C36);
  variable_set[110] = 1;
  return result;
}

static ExcelValue _common111() {
  static ExcelValue result;
  if(variable_set[111] == 1) { return result;}
  result = multiply(model_e49(),C36);
  variable_set[111] = 1;
  return result;
}

static ExcelValue _common112() {
  static ExcelValue result;
  if(variable_set[112] == 1) { return result;}
  result = multiply(model_f49(),C36);
  variable_set[112] = 1;
  return result;
}

static ExcelValue _common113() {
  static ExcelValue result;
  if(variable_set[113] == 1) { return result;}
  result = multiply(model_g49(),C36);
  variable_set[113] = 1;
  return result;
}

static ExcelValue _common114() {
  static ExcelValue result;
  if(variable_set[114] == 1) { return result;}
  result = multiply(model_h49(),C36);
  variable_set[114] = 1;
  return result;
}

static ExcelValue _common115() {
  static ExcelValue result;
  if(variable_set[115] == 1) { return result;}
  result = multiply(model_i49(),C36);
  variable_set[115] = 1;
  return result;
}

static ExcelValue _common116() {
  static ExcelValue result;
  if(variable_set[116] == 1) { return result;}
  result = add(subtract(model_c49(),C17),C35);
  variable_set[116] = 1;
  return result;
}

static ExcelValue _common117() {
  static ExcelValue result;
  if(variable_set[117] == 1) { return result;}
  result = subtract(model_c49(),C17);
  variable_set[117] = 1;
  return result;
}

static ExcelValue _common118() {
  static ExcelValue result;
  if(variable_set[118] == 1) { return result;}
  result = subtract(model_d49(),model_c49());
  variable_set[118] = 1;
  return result;
}

static ExcelValue _common119() {
  static ExcelValue result;
  if(variable_set[119] == 1) { return result;}
  result = subtract(model_e49(),model_d49());
  variable_set[119] = 1;
  return result;
}

static ExcelValue _common120() {
  static ExcelValue result;
  if(variable_set[120] == 1) { return result;}
  result = subtract(model_f49(),model_e49());
  variable_set[120] = 1;
  return result;
}

static ExcelValue _common121() {
  static ExcelValue result;
  if(variable_set[121] == 1) { return result;}
  result = subtract(model_g49(),model_f49());
  variable_set[121] = 1;
  return result;
}

static ExcelValue _common122() {
  static ExcelValue result;
  if(variable_set[122] == 1) { return result;}
  result = subtract(model_h49(),model_g49());
  variable_set[122] = 1;
  return result;
}

static ExcelValue _common123() {
  static ExcelValue result;
  if(variable_set[123] == 1) { return result;}
  result = subtract(model_i49(),model_h49());
  variable_set[123] = 1;
  return result;
}

static ExcelValue _common124() {
  static ExcelValue result;
  if(variable_set[124] == 1) { return result;}
  result = subtract(model_j49(),model_i49());
  variable_set[124] = 1;
  return result;
}

static ExcelValue _common125() {
  static ExcelValue result;
  if(variable_set[125] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_am55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_an64(),subtract(model_am74(),model_an54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_am55(),C37),C37};
  result = excel_if(more_than(model_an47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[125] = 1;
  return result;
}

static ExcelValue _common126() {
  static ExcelValue result;
  if(variable_set[126] == 1) { return result;}
  result = more_than(model_an47(),model_b8());
  variable_set[126] = 1;
  return result;
}

static ExcelValue _common127() {
  static ExcelValue result;
  if(variable_set[127] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_am55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_an64(),subtract(model_am74(),model_an54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  result = min(2, array0);
  variable_set[127] = 1;
  return result;
}

static ExcelValue _common128() {
  static ExcelValue result;
  if(variable_set[128] == 1) { return result;}
  ExcelValue array0[] = {multiply(model_am55(),C22),C6};
  result = max(2, array0);
  variable_set[128] = 1;
  return result;
}

static ExcelValue _common129() {
  static ExcelValue result;
  if(variable_set[129] == 1) { return result;}
  result = multiply(model_am55(),C22);
  variable_set[129] = 1;
  return result;
}

static ExcelValue _common130() {
  static ExcelValue result;
  if(variable_set[130] == 1) { return result;}
  ExcelValue array0[] = {model_b9(),subtract(model_an64(),subtract(model_am74(),model_an54()))};
  result = min(2, array0);
  variable_set[130] = 1;
  return result;
}

static ExcelValue _common131() {
  static ExcelValue result;
  if(variable_set[131] == 1) { return result;}
  result = subtract(model_an64(),subtract(model_am74(),model_an54()));
  variable_set[131] = 1;
  return result;
}

static ExcelValue _common132() {
  static ExcelValue result;
  if(variable_set[132] == 1) { return result;}
  result = subtract(model_am74(),model_an54());
  variable_set[132] = 1;
  return result;
}

static ExcelValue _common133() {
  static ExcelValue result;
  if(variable_set[133] == 1) { return result;}
  ExcelValue array0[] = {multiply(model_am55(),C37),C37};
  result = max(2, array0);
  variable_set[133] = 1;
  return result;
}

static ExcelValue _common134() {
  static ExcelValue result;
  if(variable_set[134] == 1) { return result;}
  result = multiply(model_am55(),C37);
  variable_set[134] = 1;
  return result;
}

static ExcelValue _common135() {
  static ExcelValue result;
  if(variable_set[135] == 1) { return result;}
  result = subtract(model_k55(),model_k54());
  variable_set[135] = 1;
  return result;
}

static ExcelValue _common136() {
  static ExcelValue result;
  if(variable_set[136] == 1) { return result;}
  result = subtract(model_l55(),model_l54());
  variable_set[136] = 1;
  return result;
}

static ExcelValue _common137() {
  static ExcelValue result;
  if(variable_set[137] == 1) { return result;}
  result = subtract(model_m55(),model_m54());
  variable_set[137] = 1;
  return result;
}

static ExcelValue _common138() {
  static ExcelValue result;
  if(variable_set[138] == 1) { return result;}
  result = subtract(model_n55(),model_n54());
  variable_set[138] = 1;
  return result;
}

static ExcelValue _common139() {
  static ExcelValue result;
  if(variable_set[139] == 1) { return result;}
  result = subtract(model_o55(),model_o54());
  variable_set[139] = 1;
  return result;
}

static ExcelValue _common140() {
  static ExcelValue result;
  if(variable_set[140] == 1) { return result;}
  result = subtract(model_p55(),model_p54());
  variable_set[140] = 1;
  return result;
}

static ExcelValue _common141() {
  static ExcelValue result;
  if(variable_set[141] == 1) { return result;}
  result = subtract(model_q55(),model_q54());
  variable_set[141] = 1;
  return result;
}

static ExcelValue _common142() {
  static ExcelValue result;
  if(variable_set[142] == 1) { return result;}
  result = subtract(model_r55(),model_r54());
  variable_set[142] = 1;
  return result;
}

static ExcelValue _common143() {
  static ExcelValue result;
  if(variable_set[143] == 1) { return result;}
  result = subtract(model_s55(),model_s54());
  variable_set[143] = 1;
  return result;
}

static ExcelValue _common144() {
  static ExcelValue result;
  if(variable_set[144] == 1) { return result;}
  result = subtract(model_t55(),model_t54());
  variable_set[144] = 1;
  return result;
}

static ExcelValue _common145() {
  static ExcelValue result;
  if(variable_set[145] == 1) { return result;}
  result = subtract(model_u55(),model_u54());
  variable_set[145] = 1;
  return result;
}

static ExcelValue _common146() {
  static ExcelValue result;
  if(variable_set[146] == 1) { return result;}
  result = subtract(model_v55(),model_v54());
  variable_set[146] = 1;
  return result;
}

static ExcelValue _common147() {
  static ExcelValue result;
  if(variable_set[147] == 1) { return result;}
  result = subtract(model_w55(),model_w54());
  variable_set[147] = 1;
  return result;
}

static ExcelValue _common148() {
  static ExcelValue result;
  if(variable_set[148] == 1) { return result;}
  result = subtract(model_x55(),model_x54());
  variable_set[148] = 1;
  return result;
}

static ExcelValue _common149() {
  static ExcelValue result;
  if(variable_set[149] == 1) { return result;}
  result = subtract(model_y55(),model_y54());
  variable_set[149] = 1;
  return result;
}

static ExcelValue _common150() {
  static ExcelValue result;
  if(variable_set[150] == 1) { return result;}
  result = subtract(model_z55(),model_z54());
  variable_set[150] = 1;
  return result;
}

static ExcelValue _common151() {
  static ExcelValue result;
  if(variable_set[151] == 1) { return result;}
  result = subtract(model_aa55(),model_aa54());
  variable_set[151] = 1;
  return result;
}

static ExcelValue _common152() {
  static ExcelValue result;
  if(variable_set[152] == 1) { return result;}
  result = subtract(model_ab55(),model_ab54());
  variable_set[152] = 1;
  return result;
}

static ExcelValue _common153() {
  static ExcelValue result;
  if(variable_set[153] == 1) { return result;}
  result = subtract(model_ac55(),model_ac54());
  variable_set[153] = 1;
  return result;
}

static ExcelValue _common154() {
  static ExcelValue result;
  if(variable_set[154] == 1) { return result;}
  result = subtract(model_ad55(),model_ad54());
  variable_set[154] = 1;
  return result;
}

static ExcelValue _common155() {
  static ExcelValue result;
  if(variable_set[155] == 1) { return result;}
  result = subtract(model_ae55(),model_ae54());
  variable_set[155] = 1;
  return result;
}

static ExcelValue _common156() {
  static ExcelValue result;
  if(variable_set[156] == 1) { return result;}
  result = subtract(model_af55(),model_af54());
  variable_set[156] = 1;
  return result;
}

static ExcelValue _common157() {
  static ExcelValue result;
  if(variable_set[157] == 1) { return result;}
  result = subtract(model_ag55(),model_ag54());
  variable_set[157] = 1;
  return result;
}

static ExcelValue _common158() {
  static ExcelValue result;
  if(variable_set[158] == 1) { return result;}
  result = subtract(model_ah55(),model_ah54());
  variable_set[158] = 1;
  return result;
}

static ExcelValue _common159() {
  static ExcelValue result;
  if(variable_set[159] == 1) { return result;}
  result = subtract(model_ai55(),model_ai54());
  variable_set[159] = 1;
  return result;
}

static ExcelValue _common160() {
  static ExcelValue result;
  if(variable_set[160] == 1) { return result;}
  result = subtract(model_aj55(),model_aj54());
  variable_set[160] = 1;
  return result;
}

static ExcelValue _common161() {
  static ExcelValue result;
  if(variable_set[161] == 1) { return result;}
  result = subtract(model_ak55(),model_ak54());
  variable_set[161] = 1;
  return result;
}

static ExcelValue _common162() {
  static ExcelValue result;
  if(variable_set[162] == 1) { return result;}
  result = subtract(model_al55(),model_al54());
  variable_set[162] = 1;
  return result;
}

static ExcelValue _common163() {
  static ExcelValue result;
  if(variable_set[163] == 1) { return result;}
  result = subtract(model_am55(),model_am54());
  variable_set[163] = 1;
  return result;
}

static ExcelValue _common164() {
  static ExcelValue result;
  if(variable_set[164] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_am55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_an64(),subtract(model_am74(),model_an54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_am55(),C37),C37};
  result = subtract(excel_if(more_than(model_an47(),model_b8()),min(2, array0),max(2, array3)),model_an54());
  variable_set[164] = 1;
  return result;
}

static ExcelValue _common165() {
  static ExcelValue result;
  if(variable_set[165] == 1) { return result;}
  result = less_than(C24,C1);
  variable_set[165] = 1;
  return result;
}

static ExcelValue _common166() {
  static ExcelValue result;
  if(variable_set[166] == 1) { return result;}
  result = less_than(model_f47(),C1);
  variable_set[166] = 1;
  return result;
}

static ExcelValue _common167() {
  static ExcelValue result;
  if(variable_set[167] == 1) { return result;}
  result = less_than(model_g47(),C1);
  variable_set[167] = 1;
  return result;
}

static ExcelValue _common168() {
  static ExcelValue result;
  if(variable_set[168] == 1) { return result;}
  result = less_than(model_h47(),C1);
  variable_set[168] = 1;
  return result;
}

static ExcelValue _common169() {
  static ExcelValue result;
  if(variable_set[169] == 1) { return result;}
  result = less_than(model_i47(),C1);
  variable_set[169] = 1;
  return result;
}

static ExcelValue _common170() {
  static ExcelValue result;
  if(variable_set[170] == 1) { return result;}
  result = less_than(model_j47(),C1);
  variable_set[170] = 1;
  return result;
}

static ExcelValue _common171() {
  static ExcelValue result;
  if(variable_set[171] == 1) { return result;}
  result = less_than(model_k47(),C1);
  variable_set[171] = 1;
  return result;
}

static ExcelValue _common172() {
  static ExcelValue result;
  if(variable_set[172] == 1) { return result;}
  result = less_than(model_l47(),C1);
  variable_set[172] = 1;
  return result;
}

static ExcelValue _common173() {
  static ExcelValue result;
  if(variable_set[173] == 1) { return result;}
  result = less_than(model_m47(),C1);
  variable_set[173] = 1;
  return result;
}

static ExcelValue _common174() {
  static ExcelValue result;
  if(variable_set[174] == 1) { return result;}
  result = less_than(model_n47(),C1);
  variable_set[174] = 1;
  return result;
}

static ExcelValue _common175() {
  static ExcelValue result;
  if(variable_set[175] == 1) { return result;}
  result = less_than(model_o47(),C1);
  variable_set[175] = 1;
  return result;
}

static ExcelValue _common176() {
  static ExcelValue result;
  if(variable_set[176] == 1) { return result;}
  result = less_than(model_p47(),C1);
  variable_set[176] = 1;
  return result;
}

static ExcelValue _common177() {
  static ExcelValue result;
  if(variable_set[177] == 1) { return result;}
  result = less_than(model_q47(),C1);
  variable_set[177] = 1;
  return result;
}

static ExcelValue _common178() {
  static ExcelValue result;
  if(variable_set[178] == 1) { return result;}
  result = less_than(model_r47(),C1);
  variable_set[178] = 1;
  return result;
}

static ExcelValue _common179() {
  static ExcelValue result;
  if(variable_set[179] == 1) { return result;}
  result = less_than(model_s47(),C1);
  variable_set[179] = 1;
  return result;
}

static ExcelValue _common180() {
  static ExcelValue result;
  if(variable_set[180] == 1) { return result;}
  result = less_than(model_t47(),C1);
  variable_set[180] = 1;
  return result;
}

static ExcelValue _common181() {
  static ExcelValue result;
  if(variable_set[181] == 1) { return result;}
  result = less_than(model_u47(),C1);
  variable_set[181] = 1;
  return result;
}

static ExcelValue _common182() {
  static ExcelValue result;
  if(variable_set[182] == 1) { return result;}
  result = less_than(model_v47(),C1);
  variable_set[182] = 1;
  return result;
}

static ExcelValue _common183() {
  static ExcelValue result;
  if(variable_set[183] == 1) { return result;}
  result = less_than(model_w47(),C1);
  variable_set[183] = 1;
  return result;
}

static ExcelValue _common184() {
  static ExcelValue result;
  if(variable_set[184] == 1) { return result;}
  result = less_than(model_x47(),C1);
  variable_set[184] = 1;
  return result;
}

static ExcelValue _common185() {
  static ExcelValue result;
  if(variable_set[185] == 1) { return result;}
  result = less_than(model_y47(),C1);
  variable_set[185] = 1;
  return result;
}

static ExcelValue _common186() {
  static ExcelValue result;
  if(variable_set[186] == 1) { return result;}
  result = less_than(model_z47(),C1);
  variable_set[186] = 1;
  return result;
}

static ExcelValue _common187() {
  static ExcelValue result;
  if(variable_set[187] == 1) { return result;}
  result = less_than(model_aa47(),C1);
  variable_set[187] = 1;
  return result;
}

static ExcelValue _common188() {
  static ExcelValue result;
  if(variable_set[188] == 1) { return result;}
  result = less_than(model_ab47(),C1);
  variable_set[188] = 1;
  return result;
}

static ExcelValue _common189() {
  static ExcelValue result;
  if(variable_set[189] == 1) { return result;}
  result = less_than(model_ac47(),C1);
  variable_set[189] = 1;
  return result;
}

static ExcelValue _common190() {
  static ExcelValue result;
  if(variable_set[190] == 1) { return result;}
  result = less_than(model_ad47(),C1);
  variable_set[190] = 1;
  return result;
}

static ExcelValue _common191() {
  static ExcelValue result;
  if(variable_set[191] == 1) { return result;}
  result = less_than(model_ae47(),C1);
  variable_set[191] = 1;
  return result;
}

static ExcelValue _common192() {
  static ExcelValue result;
  if(variable_set[192] == 1) { return result;}
  result = less_than(model_af47(),C1);
  variable_set[192] = 1;
  return result;
}

static ExcelValue _common193() {
  static ExcelValue result;
  if(variable_set[193] == 1) { return result;}
  result = less_than(model_ag47(),C1);
  variable_set[193] = 1;
  return result;
}

static ExcelValue _common194() {
  static ExcelValue result;
  if(variable_set[194] == 1) { return result;}
  result = less_than(model_ah47(),C1);
  variable_set[194] = 1;
  return result;
}

static ExcelValue _common195() {
  static ExcelValue result;
  if(variable_set[195] == 1) { return result;}
  result = less_than(model_ai47(),C1);
  variable_set[195] = 1;
  return result;
}

static ExcelValue _common196() {
  static ExcelValue result;
  if(variable_set[196] == 1) { return result;}
  result = less_than(model_aj47(),C1);
  variable_set[196] = 1;
  return result;
}

static ExcelValue _common197() {
  static ExcelValue result;
  if(variable_set[197] == 1) { return result;}
  result = less_than(model_ak47(),C1);
  variable_set[197] = 1;
  return result;
}

static ExcelValue _common198() {
  static ExcelValue result;
  if(variable_set[198] == 1) { return result;}
  result = less_than(model_al47(),C1);
  variable_set[198] = 1;
  return result;
}

static ExcelValue _common199() {
  static ExcelValue result;
  if(variable_set[199] == 1) { return result;}
  result = less_than(model_am47(),C1);
  variable_set[199] = 1;
  return result;
}

static ExcelValue _common200() {
  static ExcelValue result;
  if(variable_set[200] == 1) { return result;}
  result = less_than(model_an47(),C1);
  variable_set[200] = 1;
  return result;
}

static ExcelValue _common201() {
  static ExcelValue result;
  if(variable_set[201] == 1) { return result;}
  result = subtract(C10,C41);
  variable_set[201] = 1;
  return result;
}

static ExcelValue _common202() {
  static ExcelValue result;
  if(variable_set[202] == 1) { return result;}
  result = subtract(C25,model_c63());
  variable_set[202] = 1;
  return result;
}

static ExcelValue _common203() {
  static ExcelValue result;
  if(variable_set[203] == 1) { return result;}
  result = subtract(model_d48(),model_d63());
  variable_set[203] = 1;
  return result;
}

static ExcelValue _common204() {
  static ExcelValue result;
  if(variable_set[204] == 1) { return result;}
  result = subtract(model_e48(),model_e63());
  variable_set[204] = 1;
  return result;
}

static ExcelValue _common205() {
  static ExcelValue result;
  if(variable_set[205] == 1) { return result;}
  result = subtract(model_f48(),model_f63());
  variable_set[205] = 1;
  return result;
}

static ExcelValue _common206() {
  static ExcelValue result;
  if(variable_set[206] == 1) { return result;}
  result = subtract(model_g48(),model_g63());
  variable_set[206] = 1;
  return result;
}

static ExcelValue _common207() {
  static ExcelValue result;
  if(variable_set[207] == 1) { return result;}
  result = subtract(model_h48(),model_h63());
  variable_set[207] = 1;
  return result;
}

static ExcelValue _common208() {
  static ExcelValue result;
  if(variable_set[208] == 1) { return result;}
  result = subtract(model_i48(),model_i63());
  variable_set[208] = 1;
  return result;
}

static ExcelValue _common209() {
  static ExcelValue result;
  if(variable_set[209] == 1) { return result;}
  result = subtract(model_j48(),model_j63());
  variable_set[209] = 1;
  return result;
}

static ExcelValue _common210() {
  static ExcelValue result;
  if(variable_set[210] == 1) { return result;}
  result = subtract(model_k48(),model_k63());
  variable_set[210] = 1;
  return result;
}

static ExcelValue _common211() {
  static ExcelValue result;
  if(variable_set[211] == 1) { return result;}
  result = subtract(model_l48(),model_l63());
  variable_set[211] = 1;
  return result;
}

static ExcelValue _common212() {
  static ExcelValue result;
  if(variable_set[212] == 1) { return result;}
  result = subtract(model_m48(),model_m63());
  variable_set[212] = 1;
  return result;
}

static ExcelValue _common213() {
  static ExcelValue result;
  if(variable_set[213] == 1) { return result;}
  result = subtract(model_n48(),model_n63());
  variable_set[213] = 1;
  return result;
}

static ExcelValue _common214() {
  static ExcelValue result;
  if(variable_set[214] == 1) { return result;}
  result = subtract(model_o48(),model_o63());
  variable_set[214] = 1;
  return result;
}

static ExcelValue _common215() {
  static ExcelValue result;
  if(variable_set[215] == 1) { return result;}
  result = subtract(model_p48(),model_p63());
  variable_set[215] = 1;
  return result;
}

static ExcelValue _common216() {
  static ExcelValue result;
  if(variable_set[216] == 1) { return result;}
  result = subtract(model_q48(),model_q63());
  variable_set[216] = 1;
  return result;
}

static ExcelValue _common217() {
  static ExcelValue result;
  if(variable_set[217] == 1) { return result;}
  result = subtract(model_r48(),model_r63());
  variable_set[217] = 1;
  return result;
}

static ExcelValue _common218() {
  static ExcelValue result;
  if(variable_set[218] == 1) { return result;}
  result = subtract(model_s48(),model_s63());
  variable_set[218] = 1;
  return result;
}

static ExcelValue _common219() {
  static ExcelValue result;
  if(variable_set[219] == 1) { return result;}
  result = subtract(model_t48(),model_t63());
  variable_set[219] = 1;
  return result;
}

static ExcelValue _common220() {
  static ExcelValue result;
  if(variable_set[220] == 1) { return result;}
  result = subtract(model_u48(),model_u63());
  variable_set[220] = 1;
  return result;
}

static ExcelValue _common221() {
  static ExcelValue result;
  if(variable_set[221] == 1) { return result;}
  result = subtract(model_v48(),model_v63());
  variable_set[221] = 1;
  return result;
}

static ExcelValue _common222() {
  static ExcelValue result;
  if(variable_set[222] == 1) { return result;}
  result = subtract(model_w48(),model_w63());
  variable_set[222] = 1;
  return result;
}

static ExcelValue _common223() {
  static ExcelValue result;
  if(variable_set[223] == 1) { return result;}
  result = subtract(model_x48(),model_x63());
  variable_set[223] = 1;
  return result;
}

static ExcelValue _common224() {
  static ExcelValue result;
  if(variable_set[224] == 1) { return result;}
  result = subtract(model_y48(),model_y63());
  variable_set[224] = 1;
  return result;
}

static ExcelValue _common225() {
  static ExcelValue result;
  if(variable_set[225] == 1) { return result;}
  result = subtract(model_z48(),model_z63());
  variable_set[225] = 1;
  return result;
}

static ExcelValue _common226() {
  static ExcelValue result;
  if(variable_set[226] == 1) { return result;}
  result = subtract(model_aa48(),model_aa63());
  variable_set[226] = 1;
  return result;
}

static ExcelValue _common227() {
  static ExcelValue result;
  if(variable_set[227] == 1) { return result;}
  result = subtract(model_ab48(),model_ab63());
  variable_set[227] = 1;
  return result;
}

static ExcelValue _common228() {
  static ExcelValue result;
  if(variable_set[228] == 1) { return result;}
  result = subtract(model_ac48(),model_ac63());
  variable_set[228] = 1;
  return result;
}

static ExcelValue _common229() {
  static ExcelValue result;
  if(variable_set[229] == 1) { return result;}
  result = subtract(model_ad48(),model_ad63());
  variable_set[229] = 1;
  return result;
}

static ExcelValue _common230() {
  static ExcelValue result;
  if(variable_set[230] == 1) { return result;}
  result = subtract(model_ae48(),model_ae63());
  variable_set[230] = 1;
  return result;
}

static ExcelValue _common231() {
  static ExcelValue result;
  if(variable_set[231] == 1) { return result;}
  result = subtract(model_af48(),model_af63());
  variable_set[231] = 1;
  return result;
}

static ExcelValue _common232() {
  static ExcelValue result;
  if(variable_set[232] == 1) { return result;}
  result = subtract(model_ag48(),model_ag63());
  variable_set[232] = 1;
  return result;
}

static ExcelValue _common233() {
  static ExcelValue result;
  if(variable_set[233] == 1) { return result;}
  result = subtract(model_ah48(),model_ah63());
  variable_set[233] = 1;
  return result;
}

static ExcelValue _common234() {
  static ExcelValue result;
  if(variable_set[234] == 1) { return result;}
  result = subtract(model_ai48(),model_ai63());
  variable_set[234] = 1;
  return result;
}

static ExcelValue _common235() {
  static ExcelValue result;
  if(variable_set[235] == 1) { return result;}
  result = subtract(model_aj48(),model_aj63());
  variable_set[235] = 1;
  return result;
}

static ExcelValue _common236() {
  static ExcelValue result;
  if(variable_set[236] == 1) { return result;}
  result = subtract(model_ak48(),model_ak63());
  variable_set[236] = 1;
  return result;
}

static ExcelValue _common237() {
  static ExcelValue result;
  if(variable_set[237] == 1) { return result;}
  result = subtract(model_al48(),model_al63());
  variable_set[237] = 1;
  return result;
}

static ExcelValue _common238() {
  static ExcelValue result;
  if(variable_set[238] == 1) { return result;}
  result = subtract(model_am48(),model_am63());
  variable_set[238] = 1;
  return result;
}

static ExcelValue _common239() {
  static ExcelValue result;
  if(variable_set[239] == 1) { return result;}
  result = subtract(model_an48(),model_an63());
  variable_set[239] = 1;
  return result;
}

static ExcelValue _common240() {
  static ExcelValue result;
  if(variable_set[240] == 1) { return result;}
  result = subtract(C17,model_b75());
  variable_set[240] = 1;
  return result;
}

static ExcelValue _common241() {
  static ExcelValue result;
  if(variable_set[241] == 1) { return result;}
  result = subtract(model_c49(),model_c75());
  variable_set[241] = 1;
  return result;
}

static ExcelValue _common242() {
  static ExcelValue result;
  if(variable_set[242] == 1) { return result;}
  result = subtract(model_d49(),model_d75());
  variable_set[242] = 1;
  return result;
}

static ExcelValue _common243() {
  static ExcelValue result;
  if(variable_set[243] == 1) { return result;}
  result = subtract(model_e49(),model_e75());
  variable_set[243] = 1;
  return result;
}

static ExcelValue _common244() {
  static ExcelValue result;
  if(variable_set[244] == 1) { return result;}
  result = subtract(model_f49(),model_f75());
  variable_set[244] = 1;
  return result;
}

static ExcelValue _common245() {
  static ExcelValue result;
  if(variable_set[245] == 1) { return result;}
  result = subtract(model_g49(),model_g75());
  variable_set[245] = 1;
  return result;
}

static ExcelValue _common246() {
  static ExcelValue result;
  if(variable_set[246] == 1) { return result;}
  result = subtract(model_h49(),model_h75());
  variable_set[246] = 1;
  return result;
}

static ExcelValue _common247() {
  static ExcelValue result;
  if(variable_set[247] == 1) { return result;}
  result = subtract(model_i49(),model_i75());
  variable_set[247] = 1;
  return result;
}

static ExcelValue _common248() {
  static ExcelValue result;
  if(variable_set[248] == 1) { return result;}
  result = subtract(model_j49(),model_j75());
  variable_set[248] = 1;
  return result;
}

static ExcelValue _common249() {
  static ExcelValue result;
  if(variable_set[249] == 1) { return result;}
  result = subtract(model_k74(),model_k75());
  variable_set[249] = 1;
  return result;
}

static ExcelValue _common250() {
  static ExcelValue result;
  if(variable_set[250] == 1) { return result;}
  result = subtract(model_l74(),model_l75());
  variable_set[250] = 1;
  return result;
}

static ExcelValue _common251() {
  static ExcelValue result;
  if(variable_set[251] == 1) { return result;}
  result = subtract(model_m74(),model_m75());
  variable_set[251] = 1;
  return result;
}

static ExcelValue _common252() {
  static ExcelValue result;
  if(variable_set[252] == 1) { return result;}
  result = subtract(model_n74(),model_n75());
  variable_set[252] = 1;
  return result;
}

static ExcelValue _common253() {
  static ExcelValue result;
  if(variable_set[253] == 1) { return result;}
  result = subtract(model_o74(),model_o75());
  variable_set[253] = 1;
  return result;
}

static ExcelValue _common254() {
  static ExcelValue result;
  if(variable_set[254] == 1) { return result;}
  result = subtract(model_p74(),model_p75());
  variable_set[254] = 1;
  return result;
}

static ExcelValue _common255() {
  static ExcelValue result;
  if(variable_set[255] == 1) { return result;}
  result = subtract(model_q74(),model_q75());
  variable_set[255] = 1;
  return result;
}

static ExcelValue _common256() {
  static ExcelValue result;
  if(variable_set[256] == 1) { return result;}
  result = subtract(model_r74(),model_r75());
  variable_set[256] = 1;
  return result;
}

static ExcelValue _common257() {
  static ExcelValue result;
  if(variable_set[257] == 1) { return result;}
  result = subtract(model_s74(),model_s75());
  variable_set[257] = 1;
  return result;
}

static ExcelValue _common258() {
  static ExcelValue result;
  if(variable_set[258] == 1) { return result;}
  result = subtract(model_t74(),model_t75());
  variable_set[258] = 1;
  return result;
}

static ExcelValue _common259() {
  static ExcelValue result;
  if(variable_set[259] == 1) { return result;}
  result = subtract(model_u74(),model_u75());
  variable_set[259] = 1;
  return result;
}

static ExcelValue _common260() {
  static ExcelValue result;
  if(variable_set[260] == 1) { return result;}
  result = subtract(model_v74(),model_v75());
  variable_set[260] = 1;
  return result;
}

static ExcelValue _common261() {
  static ExcelValue result;
  if(variable_set[261] == 1) { return result;}
  result = subtract(model_w74(),model_w75());
  variable_set[261] = 1;
  return result;
}

static ExcelValue _common262() {
  static ExcelValue result;
  if(variable_set[262] == 1) { return result;}
  result = subtract(model_x74(),model_x75());
  variable_set[262] = 1;
  return result;
}

static ExcelValue _common263() {
  static ExcelValue result;
  if(variable_set[263] == 1) { return result;}
  result = subtract(model_y74(),model_y75());
  variable_set[263] = 1;
  return result;
}

static ExcelValue _common264() {
  static ExcelValue result;
  if(variable_set[264] == 1) { return result;}
  result = subtract(model_z74(),model_z75());
  variable_set[264] = 1;
  return result;
}

static ExcelValue _common265() {
  static ExcelValue result;
  if(variable_set[265] == 1) { return result;}
  result = subtract(model_aa74(),model_aa75());
  variable_set[265] = 1;
  return result;
}

static ExcelValue _common266() {
  static ExcelValue result;
  if(variable_set[266] == 1) { return result;}
  result = subtract(model_ab74(),model_ab75());
  variable_set[266] = 1;
  return result;
}

static ExcelValue _common267() {
  static ExcelValue result;
  if(variable_set[267] == 1) { return result;}
  result = subtract(model_ac74(),model_ac75());
  variable_set[267] = 1;
  return result;
}

static ExcelValue _common268() {
  static ExcelValue result;
  if(variable_set[268] == 1) { return result;}
  result = subtract(model_ad74(),model_ad75());
  variable_set[268] = 1;
  return result;
}

static ExcelValue _common269() {
  static ExcelValue result;
  if(variable_set[269] == 1) { return result;}
  result = subtract(model_ae74(),model_ae75());
  variable_set[269] = 1;
  return result;
}

static ExcelValue _common270() {
  static ExcelValue result;
  if(variable_set[270] == 1) { return result;}
  result = subtract(model_af74(),model_af75());
  variable_set[270] = 1;
  return result;
}

static ExcelValue _common271() {
  static ExcelValue result;
  if(variable_set[271] == 1) { return result;}
  result = subtract(model_ag74(),model_ag75());
  variable_set[271] = 1;
  return result;
}

static ExcelValue _common272() {
  static ExcelValue result;
  if(variable_set[272] == 1) { return result;}
  result = subtract(model_ah74(),model_ah75());
  variable_set[272] = 1;
  return result;
}

static ExcelValue _common273() {
  static ExcelValue result;
  if(variable_set[273] == 1) { return result;}
  result = subtract(model_ai74(),model_ai75());
  variable_set[273] = 1;
  return result;
}

static ExcelValue _common274() {
  static ExcelValue result;
  if(variable_set[274] == 1) { return result;}
  result = subtract(model_aj74(),model_aj75());
  variable_set[274] = 1;
  return result;
}

static ExcelValue _common275() {
  static ExcelValue result;
  if(variable_set[275] == 1) { return result;}
  result = subtract(model_ak74(),model_ak75());
  variable_set[275] = 1;
  return result;
}

static ExcelValue _common276() {
  static ExcelValue result;
  if(variable_set[276] == 1) { return result;}
  result = subtract(model_al74(),model_al75());
  variable_set[276] = 1;
  return result;
}

static ExcelValue _common277() {
  static ExcelValue result;
  if(variable_set[277] == 1) { return result;}
  result = subtract(model_am74(),model_am75());
  variable_set[277] = 1;
  return result;
}

static ExcelValue _common278() {
  static ExcelValue result;
  if(variable_set[278] == 1) { return result;}
  result = subtract(model_an74(),model_an75());
  variable_set[278] = 1;
  return result;
}

static ExcelValue _common279() {
  static ExcelValue result;
  if(variable_set[279] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_b75();
  array3[1] = model_b76();
  array3[2] = model_b77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_b75();
  array1[1] = model_b76();
  array1[2] = model_b77();
  array1[3] = subtract(C17,sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_b76(),model_b67()),subtract(C8,model_b72())))};
  ExcelValue array6[] = {model_b72(),subtract(model_b72(),multiply(divide(model_b77(),model_b68()),model_b72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),C17);
  variable_set[279] = 1;
  return result;
}

static ExcelValue _common280() {
  static ExcelValue result;
  if(variable_set[280] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_b75();
  array3[1] = model_b76();
  array3[2] = model_b77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_b75();
  array1[1] = model_b76();
  array1[2] = model_b77();
  array1[3] = subtract(C17,sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_b76(),model_b67()),subtract(C8,model_b72())))};
  ExcelValue array6[] = {model_b72(),subtract(model_b72(),multiply(divide(model_b77(),model_b68()),model_b72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[280] = 1;
  return result;
}

static ExcelValue _common281() {
  static ExcelValue result;
  if(variable_set[281] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_b75();
  array2[1] = model_b76();
  array2[2] = model_b77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_b75();
  array0[1] = model_b76();
  array0[2] = model_b77();
  array0[3] = subtract(C17,sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[281] = 1;
  return result;
}

static ExcelValue _common282() {
  static ExcelValue result;
  if(variable_set[282] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_b76(),model_b67()),subtract(C8,model_b72())))};
  ExcelValue array2[] = {model_b72(),subtract(model_b72(),multiply(divide(model_b77(),model_b68()),model_b72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[282] = 1;
  return result;
}

static ExcelValue _common283() {
  static ExcelValue result;
  if(variable_set[283] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_c75();
  array3[1] = model_c76();
  array3[2] = model_c77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_c75();
  array1[1] = model_c76();
  array1[2] = model_c77();
  array1[3] = subtract(model_c49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_c76(),model_c67()),subtract(C8,model_c72())))};
  ExcelValue array6[] = {model_c72(),subtract(model_c72(),multiply(divide(model_c77(),model_c68()),model_c72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_c49());
  variable_set[283] = 1;
  return result;
}

static ExcelValue _common284() {
  static ExcelValue result;
  if(variable_set[284] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_c75();
  array3[1] = model_c76();
  array3[2] = model_c77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_c75();
  array1[1] = model_c76();
  array1[2] = model_c77();
  array1[3] = subtract(model_c49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_c76(),model_c67()),subtract(C8,model_c72())))};
  ExcelValue array6[] = {model_c72(),subtract(model_c72(),multiply(divide(model_c77(),model_c68()),model_c72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[284] = 1;
  return result;
}

static ExcelValue _common285() {
  static ExcelValue result;
  if(variable_set[285] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_c75();
  array2[1] = model_c76();
  array2[2] = model_c77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_c75();
  array0[1] = model_c76();
  array0[2] = model_c77();
  array0[3] = subtract(model_c49(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[285] = 1;
  return result;
}

static ExcelValue _common286() {
  static ExcelValue result;
  if(variable_set[286] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_c76(),model_c67()),subtract(C8,model_c72())))};
  ExcelValue array2[] = {model_c72(),subtract(model_c72(),multiply(divide(model_c77(),model_c68()),model_c72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[286] = 1;
  return result;
}

static ExcelValue _common287() {
  static ExcelValue result;
  if(variable_set[287] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_d75();
  array3[1] = model_d76();
  array3[2] = model_d77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_d75();
  array1[1] = model_d76();
  array1[2] = model_d77();
  array1[3] = subtract(model_d49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_d76(),model_d67()),subtract(C8,model_d72())))};
  ExcelValue array6[] = {model_d72(),subtract(model_d72(),multiply(divide(model_d77(),model_d68()),model_d72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_d49());
  variable_set[287] = 1;
  return result;
}

static ExcelValue _common288() {
  static ExcelValue result;
  if(variable_set[288] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_d75();
  array3[1] = model_d76();
  array3[2] = model_d77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_d75();
  array1[1] = model_d76();
  array1[2] = model_d77();
  array1[3] = subtract(model_d49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_d76(),model_d67()),subtract(C8,model_d72())))};
  ExcelValue array6[] = {model_d72(),subtract(model_d72(),multiply(divide(model_d77(),model_d68()),model_d72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[288] = 1;
  return result;
}

static ExcelValue _common289() {
  static ExcelValue result;
  if(variable_set[289] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_d75();
  array2[1] = model_d76();
  array2[2] = model_d77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_d75();
  array0[1] = model_d76();
  array0[2] = model_d77();
  array0[3] = subtract(model_d49(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[289] = 1;
  return result;
}

static ExcelValue _common290() {
  static ExcelValue result;
  if(variable_set[290] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_d76(),model_d67()),subtract(C8,model_d72())))};
  ExcelValue array2[] = {model_d72(),subtract(model_d72(),multiply(divide(model_d77(),model_d68()),model_d72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[290] = 1;
  return result;
}

static ExcelValue _common291() {
  static ExcelValue result;
  if(variable_set[291] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_e75();
  array3[1] = model_e76();
  array3[2] = model_e77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_e75();
  array1[1] = model_e76();
  array1[2] = model_e77();
  array1[3] = subtract(model_e49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_e76(),model_e67()),subtract(C8,model_e72())))};
  ExcelValue array6[] = {model_e72(),subtract(model_e72(),multiply(divide(model_e77(),model_e68()),model_e72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_e49());
  variable_set[291] = 1;
  return result;
}

static ExcelValue _common292() {
  static ExcelValue result;
  if(variable_set[292] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_e75();
  array3[1] = model_e76();
  array3[2] = model_e77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_e75();
  array1[1] = model_e76();
  array1[2] = model_e77();
  array1[3] = subtract(model_e49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_e76(),model_e67()),subtract(C8,model_e72())))};
  ExcelValue array6[] = {model_e72(),subtract(model_e72(),multiply(divide(model_e77(),model_e68()),model_e72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[292] = 1;
  return result;
}

static ExcelValue _common293() {
  static ExcelValue result;
  if(variable_set[293] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_e75();
  array2[1] = model_e76();
  array2[2] = model_e77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_e75();
  array0[1] = model_e76();
  array0[2] = model_e77();
  array0[3] = subtract(model_e49(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[293] = 1;
  return result;
}

static ExcelValue _common294() {
  static ExcelValue result;
  if(variable_set[294] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_e76(),model_e67()),subtract(C8,model_e72())))};
  ExcelValue array2[] = {model_e72(),subtract(model_e72(),multiply(divide(model_e77(),model_e68()),model_e72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[294] = 1;
  return result;
}

static ExcelValue _common295() {
  static ExcelValue result;
  if(variable_set[295] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_f75();
  array3[1] = model_f76();
  array3[2] = model_f77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_f75();
  array1[1] = model_f76();
  array1[2] = model_f77();
  array1[3] = subtract(model_f49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_f76(),model_f67()),subtract(C8,model_f72())))};
  ExcelValue array6[] = {model_f72(),subtract(model_f72(),multiply(divide(model_f77(),model_f68()),model_f72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_f49());
  variable_set[295] = 1;
  return result;
}

static ExcelValue _common296() {
  static ExcelValue result;
  if(variable_set[296] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_f75();
  array3[1] = model_f76();
  array3[2] = model_f77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_f75();
  array1[1] = model_f76();
  array1[2] = model_f77();
  array1[3] = subtract(model_f49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_f76(),model_f67()),subtract(C8,model_f72())))};
  ExcelValue array6[] = {model_f72(),subtract(model_f72(),multiply(divide(model_f77(),model_f68()),model_f72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[296] = 1;
  return result;
}

static ExcelValue _common297() {
  static ExcelValue result;
  if(variable_set[297] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_f75();
  array2[1] = model_f76();
  array2[2] = model_f77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_f75();
  array0[1] = model_f76();
  array0[2] = model_f77();
  array0[3] = subtract(model_f49(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[297] = 1;
  return result;
}

static ExcelValue _common298() {
  static ExcelValue result;
  if(variable_set[298] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_f76(),model_f67()),subtract(C8,model_f72())))};
  ExcelValue array2[] = {model_f72(),subtract(model_f72(),multiply(divide(model_f77(),model_f68()),model_f72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[298] = 1;
  return result;
}

static ExcelValue _common299() {
  static ExcelValue result;
  if(variable_set[299] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_g75();
  array3[1] = model_g76();
  array3[2] = model_g77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_g75();
  array1[1] = model_g76();
  array1[2] = model_g77();
  array1[3] = subtract(model_g49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_g76(),model_g67()),subtract(C8,model_g72())))};
  ExcelValue array6[] = {model_g72(),subtract(model_g72(),multiply(divide(model_g77(),model_g68()),model_g72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_g49());
  variable_set[299] = 1;
  return result;
}

static ExcelValue _common300() {
  static ExcelValue result;
  if(variable_set[300] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_g75();
  array3[1] = model_g76();
  array3[2] = model_g77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_g75();
  array1[1] = model_g76();
  array1[2] = model_g77();
  array1[3] = subtract(model_g49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_g76(),model_g67()),subtract(C8,model_g72())))};
  ExcelValue array6[] = {model_g72(),subtract(model_g72(),multiply(divide(model_g77(),model_g68()),model_g72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[300] = 1;
  return result;
}

static ExcelValue _common301() {
  static ExcelValue result;
  if(variable_set[301] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_g75();
  array2[1] = model_g76();
  array2[2] = model_g77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_g75();
  array0[1] = model_g76();
  array0[2] = model_g77();
  array0[3] = subtract(model_g49(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[301] = 1;
  return result;
}

static ExcelValue _common302() {
  static ExcelValue result;
  if(variable_set[302] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_g76(),model_g67()),subtract(C8,model_g72())))};
  ExcelValue array2[] = {model_g72(),subtract(model_g72(),multiply(divide(model_g77(),model_g68()),model_g72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[302] = 1;
  return result;
}

static ExcelValue _common303() {
  static ExcelValue result;
  if(variable_set[303] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_h75();
  array3[1] = model_h76();
  array3[2] = model_h77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_h75();
  array1[1] = model_h76();
  array1[2] = model_h77();
  array1[3] = subtract(model_h49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_h76(),model_h67()),subtract(C8,model_h72())))};
  ExcelValue array6[] = {model_h72(),subtract(model_h72(),multiply(divide(model_h77(),model_h68()),model_h72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_h49());
  variable_set[303] = 1;
  return result;
}

static ExcelValue _common304() {
  static ExcelValue result;
  if(variable_set[304] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_h75();
  array3[1] = model_h76();
  array3[2] = model_h77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_h75();
  array1[1] = model_h76();
  array1[2] = model_h77();
  array1[3] = subtract(model_h49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_h76(),model_h67()),subtract(C8,model_h72())))};
  ExcelValue array6[] = {model_h72(),subtract(model_h72(),multiply(divide(model_h77(),model_h68()),model_h72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[304] = 1;
  return result;
}

static ExcelValue _common305() {
  static ExcelValue result;
  if(variable_set[305] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_h75();
  array2[1] = model_h76();
  array2[2] = model_h77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_h75();
  array0[1] = model_h76();
  array0[2] = model_h77();
  array0[3] = subtract(model_h49(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[305] = 1;
  return result;
}

static ExcelValue _common306() {
  static ExcelValue result;
  if(variable_set[306] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_h76(),model_h67()),subtract(C8,model_h72())))};
  ExcelValue array2[] = {model_h72(),subtract(model_h72(),multiply(divide(model_h77(),model_h68()),model_h72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[306] = 1;
  return result;
}

static ExcelValue _common307() {
  static ExcelValue result;
  if(variable_set[307] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_i75();
  array3[1] = model_i76();
  array3[2] = model_i77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_i75();
  array1[1] = model_i76();
  array1[2] = model_i77();
  array1[3] = subtract(model_i49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_i76(),model_i67()),subtract(C8,model_i72())))};
  ExcelValue array6[] = {model_i72(),subtract(model_i72(),multiply(divide(model_i77(),model_i68()),model_i72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_i49());
  variable_set[307] = 1;
  return result;
}

static ExcelValue _common308() {
  static ExcelValue result;
  if(variable_set[308] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_i75();
  array3[1] = model_i76();
  array3[2] = model_i77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_i75();
  array1[1] = model_i76();
  array1[2] = model_i77();
  array1[3] = subtract(model_i49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_i76(),model_i67()),subtract(C8,model_i72())))};
  ExcelValue array6[] = {model_i72(),subtract(model_i72(),multiply(divide(model_i77(),model_i68()),model_i72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[308] = 1;
  return result;
}

static ExcelValue _common309() {
  static ExcelValue result;
  if(variable_set[309] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_i75();
  array2[1] = model_i76();
  array2[2] = model_i77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_i75();
  array0[1] = model_i76();
  array0[2] = model_i77();
  array0[3] = subtract(model_i49(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[309] = 1;
  return result;
}

static ExcelValue _common310() {
  static ExcelValue result;
  if(variable_set[310] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_i76(),model_i67()),subtract(C8,model_i72())))};
  ExcelValue array2[] = {model_i72(),subtract(model_i72(),multiply(divide(model_i77(),model_i68()),model_i72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[310] = 1;
  return result;
}

static ExcelValue _common311() {
  static ExcelValue result;
  if(variable_set[311] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_j75();
  array3[1] = model_j76();
  array3[2] = model_j77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_j75();
  array1[1] = model_j76();
  array1[2] = model_j77();
  array1[3] = subtract(model_j49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_j76(),model_j67()),subtract(C8,model_j72())))};
  ExcelValue array6[] = {model_j72(),subtract(model_j72(),multiply(divide(model_j77(),model_j68()),model_j72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_j49());
  variable_set[311] = 1;
  return result;
}

static ExcelValue _common312() {
  static ExcelValue result;
  if(variable_set[312] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_j75();
  array3[1] = model_j76();
  array3[2] = model_j77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_j75();
  array1[1] = model_j76();
  array1[2] = model_j77();
  array1[3] = subtract(model_j49(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_j76(),model_j67()),subtract(C8,model_j72())))};
  ExcelValue array6[] = {model_j72(),subtract(model_j72(),multiply(divide(model_j77(),model_j68()),model_j72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[312] = 1;
  return result;
}

static ExcelValue _common313() {
  static ExcelValue result;
  if(variable_set[313] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_j75();
  array2[1] = model_j76();
  array2[2] = model_j77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_j75();
  array0[1] = model_j76();
  array0[2] = model_j77();
  array0[3] = subtract(model_j49(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[313] = 1;
  return result;
}

static ExcelValue _common314() {
  static ExcelValue result;
  if(variable_set[314] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_j76(),model_j67()),subtract(C8,model_j72())))};
  ExcelValue array2[] = {model_j72(),subtract(model_j72(),multiply(divide(model_j77(),model_j68()),model_j72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[314] = 1;
  return result;
}

static ExcelValue _common315() {
  static ExcelValue result;
  if(variable_set[315] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_k75();
  array3[1] = model_k76();
  array3[2] = model_k77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_k75();
  array1[1] = model_k76();
  array1[2] = model_k77();
  array1[3] = subtract(model_k74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_k76(),model_k67()),subtract(C8,model_k72())))};
  ExcelValue array6[] = {model_k72(),subtract(model_k72(),multiply(divide(model_k77(),model_k68()),model_k72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_k74());
  variable_set[315] = 1;
  return result;
}

static ExcelValue _common316() {
  static ExcelValue result;
  if(variable_set[316] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_k75();
  array3[1] = model_k76();
  array3[2] = model_k77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_k75();
  array1[1] = model_k76();
  array1[2] = model_k77();
  array1[3] = subtract(model_k74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_k76(),model_k67()),subtract(C8,model_k72())))};
  ExcelValue array6[] = {model_k72(),subtract(model_k72(),multiply(divide(model_k77(),model_k68()),model_k72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[316] = 1;
  return result;
}

static ExcelValue _common317() {
  static ExcelValue result;
  if(variable_set[317] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_k75();
  array2[1] = model_k76();
  array2[2] = model_k77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_k75();
  array0[1] = model_k76();
  array0[2] = model_k77();
  array0[3] = subtract(model_k74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[317] = 1;
  return result;
}

static ExcelValue _common318() {
  static ExcelValue result;
  if(variable_set[318] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_k76(),model_k67()),subtract(C8,model_k72())))};
  ExcelValue array2[] = {model_k72(),subtract(model_k72(),multiply(divide(model_k77(),model_k68()),model_k72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[318] = 1;
  return result;
}

static ExcelValue _common319() {
  static ExcelValue result;
  if(variable_set[319] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_l75();
  array3[1] = model_l76();
  array3[2] = model_l77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_l75();
  array1[1] = model_l76();
  array1[2] = model_l77();
  array1[3] = subtract(model_l74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_l76(),model_l67()),subtract(C8,model_l72())))};
  ExcelValue array6[] = {model_l72(),subtract(model_l72(),multiply(divide(model_l77(),model_l68()),model_l72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_l74());
  variable_set[319] = 1;
  return result;
}

static ExcelValue _common320() {
  static ExcelValue result;
  if(variable_set[320] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_l75();
  array3[1] = model_l76();
  array3[2] = model_l77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_l75();
  array1[1] = model_l76();
  array1[2] = model_l77();
  array1[3] = subtract(model_l74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_l76(),model_l67()),subtract(C8,model_l72())))};
  ExcelValue array6[] = {model_l72(),subtract(model_l72(),multiply(divide(model_l77(),model_l68()),model_l72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[320] = 1;
  return result;
}

static ExcelValue _common321() {
  static ExcelValue result;
  if(variable_set[321] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_l75();
  array2[1] = model_l76();
  array2[2] = model_l77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_l75();
  array0[1] = model_l76();
  array0[2] = model_l77();
  array0[3] = subtract(model_l74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[321] = 1;
  return result;
}

static ExcelValue _common322() {
  static ExcelValue result;
  if(variable_set[322] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_l76(),model_l67()),subtract(C8,model_l72())))};
  ExcelValue array2[] = {model_l72(),subtract(model_l72(),multiply(divide(model_l77(),model_l68()),model_l72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[322] = 1;
  return result;
}

static ExcelValue _common323() {
  static ExcelValue result;
  if(variable_set[323] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_m75();
  array3[1] = model_m76();
  array3[2] = model_m77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_m75();
  array1[1] = model_m76();
  array1[2] = model_m77();
  array1[3] = subtract(model_m74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_m76(),model_m67()),subtract(C8,model_m72())))};
  ExcelValue array6[] = {model_m72(),subtract(model_m72(),multiply(divide(model_m77(),model_m68()),model_m72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_m74());
  variable_set[323] = 1;
  return result;
}

static ExcelValue _common324() {
  static ExcelValue result;
  if(variable_set[324] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_m75();
  array3[1] = model_m76();
  array3[2] = model_m77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_m75();
  array1[1] = model_m76();
  array1[2] = model_m77();
  array1[3] = subtract(model_m74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_m76(),model_m67()),subtract(C8,model_m72())))};
  ExcelValue array6[] = {model_m72(),subtract(model_m72(),multiply(divide(model_m77(),model_m68()),model_m72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[324] = 1;
  return result;
}

static ExcelValue _common325() {
  static ExcelValue result;
  if(variable_set[325] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_m75();
  array2[1] = model_m76();
  array2[2] = model_m77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_m75();
  array0[1] = model_m76();
  array0[2] = model_m77();
  array0[3] = subtract(model_m74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[325] = 1;
  return result;
}

static ExcelValue _common326() {
  static ExcelValue result;
  if(variable_set[326] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_m76(),model_m67()),subtract(C8,model_m72())))};
  ExcelValue array2[] = {model_m72(),subtract(model_m72(),multiply(divide(model_m77(),model_m68()),model_m72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[326] = 1;
  return result;
}

static ExcelValue _common327() {
  static ExcelValue result;
  if(variable_set[327] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_n75();
  array3[1] = model_n76();
  array3[2] = model_n77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_n75();
  array1[1] = model_n76();
  array1[2] = model_n77();
  array1[3] = subtract(model_n74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_n76(),model_n67()),subtract(C8,model_n72())))};
  ExcelValue array6[] = {model_n72(),subtract(model_n72(),multiply(divide(model_n77(),model_n68()),model_n72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_n74());
  variable_set[327] = 1;
  return result;
}

static ExcelValue _common328() {
  static ExcelValue result;
  if(variable_set[328] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_n75();
  array3[1] = model_n76();
  array3[2] = model_n77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_n75();
  array1[1] = model_n76();
  array1[2] = model_n77();
  array1[3] = subtract(model_n74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_n76(),model_n67()),subtract(C8,model_n72())))};
  ExcelValue array6[] = {model_n72(),subtract(model_n72(),multiply(divide(model_n77(),model_n68()),model_n72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[328] = 1;
  return result;
}

static ExcelValue _common329() {
  static ExcelValue result;
  if(variable_set[329] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_n75();
  array2[1] = model_n76();
  array2[2] = model_n77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_n75();
  array0[1] = model_n76();
  array0[2] = model_n77();
  array0[3] = subtract(model_n74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[329] = 1;
  return result;
}

static ExcelValue _common330() {
  static ExcelValue result;
  if(variable_set[330] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_n76(),model_n67()),subtract(C8,model_n72())))};
  ExcelValue array2[] = {model_n72(),subtract(model_n72(),multiply(divide(model_n77(),model_n68()),model_n72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[330] = 1;
  return result;
}

static ExcelValue _common331() {
  static ExcelValue result;
  if(variable_set[331] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_o75();
  array3[1] = model_o76();
  array3[2] = model_o77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_o75();
  array1[1] = model_o76();
  array1[2] = model_o77();
  array1[3] = subtract(model_o74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_o76(),model_o67()),subtract(C8,model_o72())))};
  ExcelValue array6[] = {model_o72(),subtract(model_o72(),multiply(divide(model_o77(),model_o68()),model_o72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_o74());
  variable_set[331] = 1;
  return result;
}

static ExcelValue _common332() {
  static ExcelValue result;
  if(variable_set[332] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_o75();
  array3[1] = model_o76();
  array3[2] = model_o77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_o75();
  array1[1] = model_o76();
  array1[2] = model_o77();
  array1[3] = subtract(model_o74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_o76(),model_o67()),subtract(C8,model_o72())))};
  ExcelValue array6[] = {model_o72(),subtract(model_o72(),multiply(divide(model_o77(),model_o68()),model_o72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[332] = 1;
  return result;
}

static ExcelValue _common333() {
  static ExcelValue result;
  if(variable_set[333] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_o75();
  array2[1] = model_o76();
  array2[2] = model_o77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_o75();
  array0[1] = model_o76();
  array0[2] = model_o77();
  array0[3] = subtract(model_o74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[333] = 1;
  return result;
}

static ExcelValue _common334() {
  static ExcelValue result;
  if(variable_set[334] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_o76(),model_o67()),subtract(C8,model_o72())))};
  ExcelValue array2[] = {model_o72(),subtract(model_o72(),multiply(divide(model_o77(),model_o68()),model_o72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[334] = 1;
  return result;
}

static ExcelValue _common335() {
  static ExcelValue result;
  if(variable_set[335] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_p75();
  array3[1] = model_p76();
  array3[2] = model_p77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_p75();
  array1[1] = model_p76();
  array1[2] = model_p77();
  array1[3] = subtract(model_p74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_p76(),model_p67()),subtract(C8,model_p72())))};
  ExcelValue array6[] = {model_p72(),subtract(model_p72(),multiply(divide(model_p77(),model_p68()),model_p72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_p74());
  variable_set[335] = 1;
  return result;
}

static ExcelValue _common336() {
  static ExcelValue result;
  if(variable_set[336] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_p75();
  array3[1] = model_p76();
  array3[2] = model_p77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_p75();
  array1[1] = model_p76();
  array1[2] = model_p77();
  array1[3] = subtract(model_p74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_p76(),model_p67()),subtract(C8,model_p72())))};
  ExcelValue array6[] = {model_p72(),subtract(model_p72(),multiply(divide(model_p77(),model_p68()),model_p72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[336] = 1;
  return result;
}

static ExcelValue _common337() {
  static ExcelValue result;
  if(variable_set[337] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_p75();
  array2[1] = model_p76();
  array2[2] = model_p77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_p75();
  array0[1] = model_p76();
  array0[2] = model_p77();
  array0[3] = subtract(model_p74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[337] = 1;
  return result;
}

static ExcelValue _common338() {
  static ExcelValue result;
  if(variable_set[338] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_p76(),model_p67()),subtract(C8,model_p72())))};
  ExcelValue array2[] = {model_p72(),subtract(model_p72(),multiply(divide(model_p77(),model_p68()),model_p72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[338] = 1;
  return result;
}

static ExcelValue _common339() {
  static ExcelValue result;
  if(variable_set[339] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_q75();
  array3[1] = model_q76();
  array3[2] = model_q77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_q75();
  array1[1] = model_q76();
  array1[2] = model_q77();
  array1[3] = subtract(model_q74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_q76(),model_q67()),subtract(C8,model_q72())))};
  ExcelValue array6[] = {model_q72(),subtract(model_q72(),multiply(divide(model_q77(),model_q68()),model_q72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_q74());
  variable_set[339] = 1;
  return result;
}

static ExcelValue _common340() {
  static ExcelValue result;
  if(variable_set[340] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_q75();
  array3[1] = model_q76();
  array3[2] = model_q77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_q75();
  array1[1] = model_q76();
  array1[2] = model_q77();
  array1[3] = subtract(model_q74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_q76(),model_q67()),subtract(C8,model_q72())))};
  ExcelValue array6[] = {model_q72(),subtract(model_q72(),multiply(divide(model_q77(),model_q68()),model_q72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[340] = 1;
  return result;
}

static ExcelValue _common341() {
  static ExcelValue result;
  if(variable_set[341] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_q75();
  array2[1] = model_q76();
  array2[2] = model_q77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_q75();
  array0[1] = model_q76();
  array0[2] = model_q77();
  array0[3] = subtract(model_q74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[341] = 1;
  return result;
}

static ExcelValue _common342() {
  static ExcelValue result;
  if(variable_set[342] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_q76(),model_q67()),subtract(C8,model_q72())))};
  ExcelValue array2[] = {model_q72(),subtract(model_q72(),multiply(divide(model_q77(),model_q68()),model_q72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[342] = 1;
  return result;
}

static ExcelValue _common343() {
  static ExcelValue result;
  if(variable_set[343] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_r75();
  array3[1] = model_r76();
  array3[2] = model_r77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_r75();
  array1[1] = model_r76();
  array1[2] = model_r77();
  array1[3] = subtract(model_r74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_r76(),model_r67()),subtract(C8,model_r72())))};
  ExcelValue array6[] = {model_r72(),subtract(model_r72(),multiply(divide(model_r77(),model_r68()),model_r72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_r74());
  variable_set[343] = 1;
  return result;
}

static ExcelValue _common344() {
  static ExcelValue result;
  if(variable_set[344] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_r75();
  array3[1] = model_r76();
  array3[2] = model_r77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_r75();
  array1[1] = model_r76();
  array1[2] = model_r77();
  array1[3] = subtract(model_r74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_r76(),model_r67()),subtract(C8,model_r72())))};
  ExcelValue array6[] = {model_r72(),subtract(model_r72(),multiply(divide(model_r77(),model_r68()),model_r72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[344] = 1;
  return result;
}

static ExcelValue _common345() {
  static ExcelValue result;
  if(variable_set[345] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_r75();
  array2[1] = model_r76();
  array2[2] = model_r77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_r75();
  array0[1] = model_r76();
  array0[2] = model_r77();
  array0[3] = subtract(model_r74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[345] = 1;
  return result;
}

static ExcelValue _common346() {
  static ExcelValue result;
  if(variable_set[346] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_r76(),model_r67()),subtract(C8,model_r72())))};
  ExcelValue array2[] = {model_r72(),subtract(model_r72(),multiply(divide(model_r77(),model_r68()),model_r72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[346] = 1;
  return result;
}

static ExcelValue _common347() {
  static ExcelValue result;
  if(variable_set[347] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_s75();
  array3[1] = model_s76();
  array3[2] = model_s77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_s75();
  array1[1] = model_s76();
  array1[2] = model_s77();
  array1[3] = subtract(model_s74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_s76(),model_s67()),subtract(C8,model_s72())))};
  ExcelValue array6[] = {model_s72(),subtract(model_s72(),multiply(divide(model_s77(),model_s68()),model_s72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_s74());
  variable_set[347] = 1;
  return result;
}

static ExcelValue _common348() {
  static ExcelValue result;
  if(variable_set[348] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_s75();
  array3[1] = model_s76();
  array3[2] = model_s77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_s75();
  array1[1] = model_s76();
  array1[2] = model_s77();
  array1[3] = subtract(model_s74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_s76(),model_s67()),subtract(C8,model_s72())))};
  ExcelValue array6[] = {model_s72(),subtract(model_s72(),multiply(divide(model_s77(),model_s68()),model_s72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[348] = 1;
  return result;
}

static ExcelValue _common349() {
  static ExcelValue result;
  if(variable_set[349] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_s75();
  array2[1] = model_s76();
  array2[2] = model_s77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_s75();
  array0[1] = model_s76();
  array0[2] = model_s77();
  array0[3] = subtract(model_s74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[349] = 1;
  return result;
}

static ExcelValue _common350() {
  static ExcelValue result;
  if(variable_set[350] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_s76(),model_s67()),subtract(C8,model_s72())))};
  ExcelValue array2[] = {model_s72(),subtract(model_s72(),multiply(divide(model_s77(),model_s68()),model_s72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[350] = 1;
  return result;
}

static ExcelValue _common351() {
  static ExcelValue result;
  if(variable_set[351] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_t75();
  array3[1] = model_t76();
  array3[2] = model_t77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_t75();
  array1[1] = model_t76();
  array1[2] = model_t77();
  array1[3] = subtract(model_t74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_t76(),model_t67()),subtract(C8,model_t72())))};
  ExcelValue array6[] = {model_t72(),subtract(model_t72(),multiply(divide(model_t77(),model_t68()),model_t72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_t74());
  variable_set[351] = 1;
  return result;
}

static ExcelValue _common352() {
  static ExcelValue result;
  if(variable_set[352] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_t75();
  array3[1] = model_t76();
  array3[2] = model_t77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_t75();
  array1[1] = model_t76();
  array1[2] = model_t77();
  array1[3] = subtract(model_t74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_t76(),model_t67()),subtract(C8,model_t72())))};
  ExcelValue array6[] = {model_t72(),subtract(model_t72(),multiply(divide(model_t77(),model_t68()),model_t72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[352] = 1;
  return result;
}

static ExcelValue _common353() {
  static ExcelValue result;
  if(variable_set[353] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_t75();
  array2[1] = model_t76();
  array2[2] = model_t77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_t75();
  array0[1] = model_t76();
  array0[2] = model_t77();
  array0[3] = subtract(model_t74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[353] = 1;
  return result;
}

static ExcelValue _common354() {
  static ExcelValue result;
  if(variable_set[354] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_t76(),model_t67()),subtract(C8,model_t72())))};
  ExcelValue array2[] = {model_t72(),subtract(model_t72(),multiply(divide(model_t77(),model_t68()),model_t72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[354] = 1;
  return result;
}

static ExcelValue _common355() {
  static ExcelValue result;
  if(variable_set[355] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_u75();
  array3[1] = model_u76();
  array3[2] = model_u77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_u75();
  array1[1] = model_u76();
  array1[2] = model_u77();
  array1[3] = subtract(model_u74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_u76(),model_u67()),subtract(C8,model_u72())))};
  ExcelValue array6[] = {model_u72(),subtract(model_u72(),multiply(divide(model_u77(),model_u68()),model_u72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_u74());
  variable_set[355] = 1;
  return result;
}

static ExcelValue _common356() {
  static ExcelValue result;
  if(variable_set[356] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_u75();
  array3[1] = model_u76();
  array3[2] = model_u77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_u75();
  array1[1] = model_u76();
  array1[2] = model_u77();
  array1[3] = subtract(model_u74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_u76(),model_u67()),subtract(C8,model_u72())))};
  ExcelValue array6[] = {model_u72(),subtract(model_u72(),multiply(divide(model_u77(),model_u68()),model_u72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[356] = 1;
  return result;
}

static ExcelValue _common357() {
  static ExcelValue result;
  if(variable_set[357] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_u75();
  array2[1] = model_u76();
  array2[2] = model_u77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_u75();
  array0[1] = model_u76();
  array0[2] = model_u77();
  array0[3] = subtract(model_u74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[357] = 1;
  return result;
}

static ExcelValue _common358() {
  static ExcelValue result;
  if(variable_set[358] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_u76(),model_u67()),subtract(C8,model_u72())))};
  ExcelValue array2[] = {model_u72(),subtract(model_u72(),multiply(divide(model_u77(),model_u68()),model_u72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[358] = 1;
  return result;
}

static ExcelValue _common359() {
  static ExcelValue result;
  if(variable_set[359] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_v75();
  array3[1] = model_v76();
  array3[2] = model_v77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_v75();
  array1[1] = model_v76();
  array1[2] = model_v77();
  array1[3] = subtract(model_v74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_v76(),model_v67()),subtract(C8,model_v72())))};
  ExcelValue array6[] = {model_v72(),subtract(model_v72(),multiply(divide(model_v77(),model_v68()),model_v72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_v74());
  variable_set[359] = 1;
  return result;
}

static ExcelValue _common360() {
  static ExcelValue result;
  if(variable_set[360] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_v75();
  array3[1] = model_v76();
  array3[2] = model_v77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_v75();
  array1[1] = model_v76();
  array1[2] = model_v77();
  array1[3] = subtract(model_v74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_v76(),model_v67()),subtract(C8,model_v72())))};
  ExcelValue array6[] = {model_v72(),subtract(model_v72(),multiply(divide(model_v77(),model_v68()),model_v72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[360] = 1;
  return result;
}

static ExcelValue _common361() {
  static ExcelValue result;
  if(variable_set[361] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_v75();
  array2[1] = model_v76();
  array2[2] = model_v77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_v75();
  array0[1] = model_v76();
  array0[2] = model_v77();
  array0[3] = subtract(model_v74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[361] = 1;
  return result;
}

static ExcelValue _common362() {
  static ExcelValue result;
  if(variable_set[362] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_v76(),model_v67()),subtract(C8,model_v72())))};
  ExcelValue array2[] = {model_v72(),subtract(model_v72(),multiply(divide(model_v77(),model_v68()),model_v72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[362] = 1;
  return result;
}

static ExcelValue _common363() {
  static ExcelValue result;
  if(variable_set[363] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_w75();
  array3[1] = model_w76();
  array3[2] = model_w77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_w75();
  array1[1] = model_w76();
  array1[2] = model_w77();
  array1[3] = subtract(model_w74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_w76(),model_w67()),subtract(C8,model_w72())))};
  ExcelValue array6[] = {model_w72(),subtract(model_w72(),multiply(divide(model_w77(),model_w68()),model_w72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_w74());
  variable_set[363] = 1;
  return result;
}

static ExcelValue _common364() {
  static ExcelValue result;
  if(variable_set[364] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_w75();
  array3[1] = model_w76();
  array3[2] = model_w77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_w75();
  array1[1] = model_w76();
  array1[2] = model_w77();
  array1[3] = subtract(model_w74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_w76(),model_w67()),subtract(C8,model_w72())))};
  ExcelValue array6[] = {model_w72(),subtract(model_w72(),multiply(divide(model_w77(),model_w68()),model_w72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[364] = 1;
  return result;
}

static ExcelValue _common365() {
  static ExcelValue result;
  if(variable_set[365] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_w75();
  array2[1] = model_w76();
  array2[2] = model_w77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_w75();
  array0[1] = model_w76();
  array0[2] = model_w77();
  array0[3] = subtract(model_w74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[365] = 1;
  return result;
}

static ExcelValue _common366() {
  static ExcelValue result;
  if(variable_set[366] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_w76(),model_w67()),subtract(C8,model_w72())))};
  ExcelValue array2[] = {model_w72(),subtract(model_w72(),multiply(divide(model_w77(),model_w68()),model_w72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[366] = 1;
  return result;
}

static ExcelValue _common367() {
  static ExcelValue result;
  if(variable_set[367] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_x75();
  array3[1] = model_x76();
  array3[2] = model_x77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_x75();
  array1[1] = model_x76();
  array1[2] = model_x77();
  array1[3] = subtract(model_x74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_x76(),model_x67()),subtract(C8,model_x72())))};
  ExcelValue array6[] = {model_x72(),subtract(model_x72(),multiply(divide(model_x77(),model_x68()),model_x72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_x74());
  variable_set[367] = 1;
  return result;
}

static ExcelValue _common368() {
  static ExcelValue result;
  if(variable_set[368] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_x75();
  array3[1] = model_x76();
  array3[2] = model_x77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_x75();
  array1[1] = model_x76();
  array1[2] = model_x77();
  array1[3] = subtract(model_x74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_x76(),model_x67()),subtract(C8,model_x72())))};
  ExcelValue array6[] = {model_x72(),subtract(model_x72(),multiply(divide(model_x77(),model_x68()),model_x72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[368] = 1;
  return result;
}

static ExcelValue _common369() {
  static ExcelValue result;
  if(variable_set[369] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_x75();
  array2[1] = model_x76();
  array2[2] = model_x77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_x75();
  array0[1] = model_x76();
  array0[2] = model_x77();
  array0[3] = subtract(model_x74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[369] = 1;
  return result;
}

static ExcelValue _common370() {
  static ExcelValue result;
  if(variable_set[370] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_x76(),model_x67()),subtract(C8,model_x72())))};
  ExcelValue array2[] = {model_x72(),subtract(model_x72(),multiply(divide(model_x77(),model_x68()),model_x72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[370] = 1;
  return result;
}

static ExcelValue _common371() {
  static ExcelValue result;
  if(variable_set[371] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_y75();
  array3[1] = model_y76();
  array3[2] = model_y77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_y75();
  array1[1] = model_y76();
  array1[2] = model_y77();
  array1[3] = subtract(model_y74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_y76(),model_y67()),subtract(C8,model_y72())))};
  ExcelValue array6[] = {model_y72(),subtract(model_y72(),multiply(divide(model_y77(),model_y68()),model_y72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_y74());
  variable_set[371] = 1;
  return result;
}

static ExcelValue _common372() {
  static ExcelValue result;
  if(variable_set[372] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_y75();
  array3[1] = model_y76();
  array3[2] = model_y77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_y75();
  array1[1] = model_y76();
  array1[2] = model_y77();
  array1[3] = subtract(model_y74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_y76(),model_y67()),subtract(C8,model_y72())))};
  ExcelValue array6[] = {model_y72(),subtract(model_y72(),multiply(divide(model_y77(),model_y68()),model_y72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[372] = 1;
  return result;
}

static ExcelValue _common373() {
  static ExcelValue result;
  if(variable_set[373] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_y75();
  array2[1] = model_y76();
  array2[2] = model_y77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_y75();
  array0[1] = model_y76();
  array0[2] = model_y77();
  array0[3] = subtract(model_y74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[373] = 1;
  return result;
}

static ExcelValue _common374() {
  static ExcelValue result;
  if(variable_set[374] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_y76(),model_y67()),subtract(C8,model_y72())))};
  ExcelValue array2[] = {model_y72(),subtract(model_y72(),multiply(divide(model_y77(),model_y68()),model_y72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[374] = 1;
  return result;
}

static ExcelValue _common375() {
  static ExcelValue result;
  if(variable_set[375] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_z75();
  array3[1] = model_z76();
  array3[2] = model_z77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_z75();
  array1[1] = model_z76();
  array1[2] = model_z77();
  array1[3] = subtract(model_z74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_z76(),model_z67()),subtract(C8,model_z72())))};
  ExcelValue array6[] = {model_z72(),subtract(model_z72(),multiply(divide(model_z77(),model_z68()),model_z72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_z74());
  variable_set[375] = 1;
  return result;
}

static ExcelValue _common376() {
  static ExcelValue result;
  if(variable_set[376] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_z75();
  array3[1] = model_z76();
  array3[2] = model_z77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_z75();
  array1[1] = model_z76();
  array1[2] = model_z77();
  array1[3] = subtract(model_z74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_z76(),model_z67()),subtract(C8,model_z72())))};
  ExcelValue array6[] = {model_z72(),subtract(model_z72(),multiply(divide(model_z77(),model_z68()),model_z72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[376] = 1;
  return result;
}

static ExcelValue _common377() {
  static ExcelValue result;
  if(variable_set[377] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_z75();
  array2[1] = model_z76();
  array2[2] = model_z77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_z75();
  array0[1] = model_z76();
  array0[2] = model_z77();
  array0[3] = subtract(model_z74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[377] = 1;
  return result;
}

static ExcelValue _common378() {
  static ExcelValue result;
  if(variable_set[378] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_z76(),model_z67()),subtract(C8,model_z72())))};
  ExcelValue array2[] = {model_z72(),subtract(model_z72(),multiply(divide(model_z77(),model_z68()),model_z72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[378] = 1;
  return result;
}

static ExcelValue _common379() {
  static ExcelValue result;
  if(variable_set[379] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_aa75();
  array3[1] = model_aa76();
  array3[2] = model_aa77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_aa75();
  array1[1] = model_aa76();
  array1[2] = model_aa77();
  array1[3] = subtract(model_aa74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_aa76(),model_aa67()),subtract(C8,model_aa72())))};
  ExcelValue array6[] = {model_aa72(),subtract(model_aa72(),multiply(divide(model_aa77(),model_aa68()),model_aa72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_aa74());
  variable_set[379] = 1;
  return result;
}

static ExcelValue _common380() {
  static ExcelValue result;
  if(variable_set[380] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_aa75();
  array3[1] = model_aa76();
  array3[2] = model_aa77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_aa75();
  array1[1] = model_aa76();
  array1[2] = model_aa77();
  array1[3] = subtract(model_aa74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_aa76(),model_aa67()),subtract(C8,model_aa72())))};
  ExcelValue array6[] = {model_aa72(),subtract(model_aa72(),multiply(divide(model_aa77(),model_aa68()),model_aa72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[380] = 1;
  return result;
}

static ExcelValue _common381() {
  static ExcelValue result;
  if(variable_set[381] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_aa75();
  array2[1] = model_aa76();
  array2[2] = model_aa77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_aa75();
  array0[1] = model_aa76();
  array0[2] = model_aa77();
  array0[3] = subtract(model_aa74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[381] = 1;
  return result;
}

static ExcelValue _common382() {
  static ExcelValue result;
  if(variable_set[382] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_aa76(),model_aa67()),subtract(C8,model_aa72())))};
  ExcelValue array2[] = {model_aa72(),subtract(model_aa72(),multiply(divide(model_aa77(),model_aa68()),model_aa72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[382] = 1;
  return result;
}

static ExcelValue _common383() {
  static ExcelValue result;
  if(variable_set[383] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ab75();
  array3[1] = model_ab76();
  array3[2] = model_ab77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ab75();
  array1[1] = model_ab76();
  array1[2] = model_ab77();
  array1[3] = subtract(model_ab74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ab76(),model_ab67()),subtract(C8,model_ab72())))};
  ExcelValue array6[] = {model_ab72(),subtract(model_ab72(),multiply(divide(model_ab77(),model_ab68()),model_ab72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_ab74());
  variable_set[383] = 1;
  return result;
}

static ExcelValue _common384() {
  static ExcelValue result;
  if(variable_set[384] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ab75();
  array3[1] = model_ab76();
  array3[2] = model_ab77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ab75();
  array1[1] = model_ab76();
  array1[2] = model_ab77();
  array1[3] = subtract(model_ab74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ab76(),model_ab67()),subtract(C8,model_ab72())))};
  ExcelValue array6[] = {model_ab72(),subtract(model_ab72(),multiply(divide(model_ab77(),model_ab68()),model_ab72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[384] = 1;
  return result;
}

static ExcelValue _common385() {
  static ExcelValue result;
  if(variable_set[385] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_ab75();
  array2[1] = model_ab76();
  array2[2] = model_ab77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_ab75();
  array0[1] = model_ab76();
  array0[2] = model_ab77();
  array0[3] = subtract(model_ab74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[385] = 1;
  return result;
}

static ExcelValue _common386() {
  static ExcelValue result;
  if(variable_set[386] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_ab76(),model_ab67()),subtract(C8,model_ab72())))};
  ExcelValue array2[] = {model_ab72(),subtract(model_ab72(),multiply(divide(model_ab77(),model_ab68()),model_ab72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[386] = 1;
  return result;
}

static ExcelValue _common387() {
  static ExcelValue result;
  if(variable_set[387] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ac75();
  array3[1] = model_ac76();
  array3[2] = model_ac77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ac75();
  array1[1] = model_ac76();
  array1[2] = model_ac77();
  array1[3] = subtract(model_ac74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ac76(),model_ac67()),subtract(C8,model_ac72())))};
  ExcelValue array6[] = {model_ac72(),subtract(model_ac72(),multiply(divide(model_ac77(),model_ac68()),model_ac72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_ac74());
  variable_set[387] = 1;
  return result;
}

static ExcelValue _common388() {
  static ExcelValue result;
  if(variable_set[388] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ac75();
  array3[1] = model_ac76();
  array3[2] = model_ac77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ac75();
  array1[1] = model_ac76();
  array1[2] = model_ac77();
  array1[3] = subtract(model_ac74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ac76(),model_ac67()),subtract(C8,model_ac72())))};
  ExcelValue array6[] = {model_ac72(),subtract(model_ac72(),multiply(divide(model_ac77(),model_ac68()),model_ac72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[388] = 1;
  return result;
}

static ExcelValue _common389() {
  static ExcelValue result;
  if(variable_set[389] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_ac75();
  array2[1] = model_ac76();
  array2[2] = model_ac77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_ac75();
  array0[1] = model_ac76();
  array0[2] = model_ac77();
  array0[3] = subtract(model_ac74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[389] = 1;
  return result;
}

static ExcelValue _common390() {
  static ExcelValue result;
  if(variable_set[390] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_ac76(),model_ac67()),subtract(C8,model_ac72())))};
  ExcelValue array2[] = {model_ac72(),subtract(model_ac72(),multiply(divide(model_ac77(),model_ac68()),model_ac72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[390] = 1;
  return result;
}

static ExcelValue _common391() {
  static ExcelValue result;
  if(variable_set[391] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ad75();
  array3[1] = model_ad76();
  array3[2] = model_ad77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ad75();
  array1[1] = model_ad76();
  array1[2] = model_ad77();
  array1[3] = subtract(model_ad74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ad76(),model_ad67()),subtract(C8,model_ad72())))};
  ExcelValue array6[] = {model_ad72(),subtract(model_ad72(),multiply(divide(model_ad77(),model_ad68()),model_ad72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_ad74());
  variable_set[391] = 1;
  return result;
}

static ExcelValue _common392() {
  static ExcelValue result;
  if(variable_set[392] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ad75();
  array3[1] = model_ad76();
  array3[2] = model_ad77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ad75();
  array1[1] = model_ad76();
  array1[2] = model_ad77();
  array1[3] = subtract(model_ad74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ad76(),model_ad67()),subtract(C8,model_ad72())))};
  ExcelValue array6[] = {model_ad72(),subtract(model_ad72(),multiply(divide(model_ad77(),model_ad68()),model_ad72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[392] = 1;
  return result;
}

static ExcelValue _common393() {
  static ExcelValue result;
  if(variable_set[393] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_ad75();
  array2[1] = model_ad76();
  array2[2] = model_ad77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_ad75();
  array0[1] = model_ad76();
  array0[2] = model_ad77();
  array0[3] = subtract(model_ad74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[393] = 1;
  return result;
}

static ExcelValue _common394() {
  static ExcelValue result;
  if(variable_set[394] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_ad76(),model_ad67()),subtract(C8,model_ad72())))};
  ExcelValue array2[] = {model_ad72(),subtract(model_ad72(),multiply(divide(model_ad77(),model_ad68()),model_ad72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[394] = 1;
  return result;
}

static ExcelValue _common395() {
  static ExcelValue result;
  if(variable_set[395] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ae75();
  array3[1] = model_ae76();
  array3[2] = model_ae77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ae75();
  array1[1] = model_ae76();
  array1[2] = model_ae77();
  array1[3] = subtract(model_ae74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ae76(),model_ae67()),subtract(C8,model_ae72())))};
  ExcelValue array6[] = {model_ae72(),subtract(model_ae72(),multiply(divide(model_ae77(),model_ae68()),model_ae72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_ae74());
  variable_set[395] = 1;
  return result;
}

static ExcelValue _common396() {
  static ExcelValue result;
  if(variable_set[396] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ae75();
  array3[1] = model_ae76();
  array3[2] = model_ae77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ae75();
  array1[1] = model_ae76();
  array1[2] = model_ae77();
  array1[3] = subtract(model_ae74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ae76(),model_ae67()),subtract(C8,model_ae72())))};
  ExcelValue array6[] = {model_ae72(),subtract(model_ae72(),multiply(divide(model_ae77(),model_ae68()),model_ae72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[396] = 1;
  return result;
}

static ExcelValue _common397() {
  static ExcelValue result;
  if(variable_set[397] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_ae75();
  array2[1] = model_ae76();
  array2[2] = model_ae77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_ae75();
  array0[1] = model_ae76();
  array0[2] = model_ae77();
  array0[3] = subtract(model_ae74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[397] = 1;
  return result;
}

static ExcelValue _common398() {
  static ExcelValue result;
  if(variable_set[398] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_ae76(),model_ae67()),subtract(C8,model_ae72())))};
  ExcelValue array2[] = {model_ae72(),subtract(model_ae72(),multiply(divide(model_ae77(),model_ae68()),model_ae72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[398] = 1;
  return result;
}

static ExcelValue _common399() {
  static ExcelValue result;
  if(variable_set[399] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_af75();
  array3[1] = model_af76();
  array3[2] = model_af77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_af75();
  array1[1] = model_af76();
  array1[2] = model_af77();
  array1[3] = subtract(model_af74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_af76(),model_af67()),subtract(C8,model_af72())))};
  ExcelValue array6[] = {model_af72(),subtract(model_af72(),multiply(divide(model_af77(),model_af68()),model_af72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_af74());
  variable_set[399] = 1;
  return result;
}

static ExcelValue _common400() {
  static ExcelValue result;
  if(variable_set[400] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_af75();
  array3[1] = model_af76();
  array3[2] = model_af77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_af75();
  array1[1] = model_af76();
  array1[2] = model_af77();
  array1[3] = subtract(model_af74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_af76(),model_af67()),subtract(C8,model_af72())))};
  ExcelValue array6[] = {model_af72(),subtract(model_af72(),multiply(divide(model_af77(),model_af68()),model_af72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[400] = 1;
  return result;
}

static ExcelValue _common401() {
  static ExcelValue result;
  if(variable_set[401] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_af75();
  array2[1] = model_af76();
  array2[2] = model_af77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_af75();
  array0[1] = model_af76();
  array0[2] = model_af77();
  array0[3] = subtract(model_af74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[401] = 1;
  return result;
}

static ExcelValue _common402() {
  static ExcelValue result;
  if(variable_set[402] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_af76(),model_af67()),subtract(C8,model_af72())))};
  ExcelValue array2[] = {model_af72(),subtract(model_af72(),multiply(divide(model_af77(),model_af68()),model_af72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[402] = 1;
  return result;
}

static ExcelValue _common403() {
  static ExcelValue result;
  if(variable_set[403] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ag75();
  array3[1] = model_ag76();
  array3[2] = model_ag77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ag75();
  array1[1] = model_ag76();
  array1[2] = model_ag77();
  array1[3] = subtract(model_ag74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ag76(),model_ag67()),subtract(C8,model_ag72())))};
  ExcelValue array6[] = {model_ag72(),subtract(model_ag72(),multiply(divide(model_ag77(),model_ag68()),model_ag72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_ag74());
  variable_set[403] = 1;
  return result;
}

static ExcelValue _common404() {
  static ExcelValue result;
  if(variable_set[404] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ag75();
  array3[1] = model_ag76();
  array3[2] = model_ag77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ag75();
  array1[1] = model_ag76();
  array1[2] = model_ag77();
  array1[3] = subtract(model_ag74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ag76(),model_ag67()),subtract(C8,model_ag72())))};
  ExcelValue array6[] = {model_ag72(),subtract(model_ag72(),multiply(divide(model_ag77(),model_ag68()),model_ag72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[404] = 1;
  return result;
}

static ExcelValue _common405() {
  static ExcelValue result;
  if(variable_set[405] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_ag75();
  array2[1] = model_ag76();
  array2[2] = model_ag77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_ag75();
  array0[1] = model_ag76();
  array0[2] = model_ag77();
  array0[3] = subtract(model_ag74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[405] = 1;
  return result;
}

static ExcelValue _common406() {
  static ExcelValue result;
  if(variable_set[406] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_ag76(),model_ag67()),subtract(C8,model_ag72())))};
  ExcelValue array2[] = {model_ag72(),subtract(model_ag72(),multiply(divide(model_ag77(),model_ag68()),model_ag72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[406] = 1;
  return result;
}

static ExcelValue _common407() {
  static ExcelValue result;
  if(variable_set[407] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ah75();
  array3[1] = model_ah76();
  array3[2] = model_ah77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ah75();
  array1[1] = model_ah76();
  array1[2] = model_ah77();
  array1[3] = subtract(model_ah74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ah76(),model_ah67()),subtract(C8,model_ah72())))};
  ExcelValue array6[] = {model_ah72(),subtract(model_ah72(),multiply(divide(model_ah77(),model_ah68()),model_ah72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_ah74());
  variable_set[407] = 1;
  return result;
}

static ExcelValue _common408() {
  static ExcelValue result;
  if(variable_set[408] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ah75();
  array3[1] = model_ah76();
  array3[2] = model_ah77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ah75();
  array1[1] = model_ah76();
  array1[2] = model_ah77();
  array1[3] = subtract(model_ah74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ah76(),model_ah67()),subtract(C8,model_ah72())))};
  ExcelValue array6[] = {model_ah72(),subtract(model_ah72(),multiply(divide(model_ah77(),model_ah68()),model_ah72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[408] = 1;
  return result;
}

static ExcelValue _common409() {
  static ExcelValue result;
  if(variable_set[409] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_ah75();
  array2[1] = model_ah76();
  array2[2] = model_ah77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_ah75();
  array0[1] = model_ah76();
  array0[2] = model_ah77();
  array0[3] = subtract(model_ah74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[409] = 1;
  return result;
}

static ExcelValue _common410() {
  static ExcelValue result;
  if(variable_set[410] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_ah76(),model_ah67()),subtract(C8,model_ah72())))};
  ExcelValue array2[] = {model_ah72(),subtract(model_ah72(),multiply(divide(model_ah77(),model_ah68()),model_ah72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[410] = 1;
  return result;
}

static ExcelValue _common411() {
  static ExcelValue result;
  if(variable_set[411] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ai75();
  array3[1] = model_ai76();
  array3[2] = model_ai77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ai75();
  array1[1] = model_ai76();
  array1[2] = model_ai77();
  array1[3] = subtract(model_ai74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ai76(),model_ai67()),subtract(C8,model_ai72())))};
  ExcelValue array6[] = {model_ai72(),subtract(model_ai72(),multiply(divide(model_ai77(),model_ai68()),model_ai72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_ai74());
  variable_set[411] = 1;
  return result;
}

static ExcelValue _common412() {
  static ExcelValue result;
  if(variable_set[412] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ai75();
  array3[1] = model_ai76();
  array3[2] = model_ai77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ai75();
  array1[1] = model_ai76();
  array1[2] = model_ai77();
  array1[3] = subtract(model_ai74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ai76(),model_ai67()),subtract(C8,model_ai72())))};
  ExcelValue array6[] = {model_ai72(),subtract(model_ai72(),multiply(divide(model_ai77(),model_ai68()),model_ai72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[412] = 1;
  return result;
}

static ExcelValue _common413() {
  static ExcelValue result;
  if(variable_set[413] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_ai75();
  array2[1] = model_ai76();
  array2[2] = model_ai77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_ai75();
  array0[1] = model_ai76();
  array0[2] = model_ai77();
  array0[3] = subtract(model_ai74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[413] = 1;
  return result;
}

static ExcelValue _common414() {
  static ExcelValue result;
  if(variable_set[414] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_ai76(),model_ai67()),subtract(C8,model_ai72())))};
  ExcelValue array2[] = {model_ai72(),subtract(model_ai72(),multiply(divide(model_ai77(),model_ai68()),model_ai72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[414] = 1;
  return result;
}

static ExcelValue _common415() {
  static ExcelValue result;
  if(variable_set[415] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_aj75();
  array3[1] = model_aj76();
  array3[2] = model_aj77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_aj75();
  array1[1] = model_aj76();
  array1[2] = model_aj77();
  array1[3] = subtract(model_aj74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_aj76(),model_aj67()),subtract(C8,model_aj72())))};
  ExcelValue array6[] = {model_aj72(),subtract(model_aj72(),multiply(divide(model_aj77(),model_aj68()),model_aj72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_aj74());
  variable_set[415] = 1;
  return result;
}

static ExcelValue _common416() {
  static ExcelValue result;
  if(variable_set[416] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_aj75();
  array3[1] = model_aj76();
  array3[2] = model_aj77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_aj75();
  array1[1] = model_aj76();
  array1[2] = model_aj77();
  array1[3] = subtract(model_aj74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_aj76(),model_aj67()),subtract(C8,model_aj72())))};
  ExcelValue array6[] = {model_aj72(),subtract(model_aj72(),multiply(divide(model_aj77(),model_aj68()),model_aj72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[416] = 1;
  return result;
}

static ExcelValue _common417() {
  static ExcelValue result;
  if(variable_set[417] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_aj75();
  array2[1] = model_aj76();
  array2[2] = model_aj77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_aj75();
  array0[1] = model_aj76();
  array0[2] = model_aj77();
  array0[3] = subtract(model_aj74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[417] = 1;
  return result;
}

static ExcelValue _common418() {
  static ExcelValue result;
  if(variable_set[418] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_aj76(),model_aj67()),subtract(C8,model_aj72())))};
  ExcelValue array2[] = {model_aj72(),subtract(model_aj72(),multiply(divide(model_aj77(),model_aj68()),model_aj72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[418] = 1;
  return result;
}

static ExcelValue _common419() {
  static ExcelValue result;
  if(variable_set[419] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ak75();
  array3[1] = model_ak76();
  array3[2] = model_ak77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ak75();
  array1[1] = model_ak76();
  array1[2] = model_ak77();
  array1[3] = subtract(model_ak74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ak76(),model_ak67()),subtract(C8,model_ak72())))};
  ExcelValue array6[] = {model_ak72(),subtract(model_ak72(),multiply(divide(model_ak77(),model_ak68()),model_ak72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_ak74());
  variable_set[419] = 1;
  return result;
}

static ExcelValue _common420() {
  static ExcelValue result;
  if(variable_set[420] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_ak75();
  array3[1] = model_ak76();
  array3[2] = model_ak77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_ak75();
  array1[1] = model_ak76();
  array1[2] = model_ak77();
  array1[3] = subtract(model_ak74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_ak76(),model_ak67()),subtract(C8,model_ak72())))};
  ExcelValue array6[] = {model_ak72(),subtract(model_ak72(),multiply(divide(model_ak77(),model_ak68()),model_ak72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[420] = 1;
  return result;
}

static ExcelValue _common421() {
  static ExcelValue result;
  if(variable_set[421] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_ak75();
  array2[1] = model_ak76();
  array2[2] = model_ak77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_ak75();
  array0[1] = model_ak76();
  array0[2] = model_ak77();
  array0[3] = subtract(model_ak74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[421] = 1;
  return result;
}

static ExcelValue _common422() {
  static ExcelValue result;
  if(variable_set[422] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_ak76(),model_ak67()),subtract(C8,model_ak72())))};
  ExcelValue array2[] = {model_ak72(),subtract(model_ak72(),multiply(divide(model_ak77(),model_ak68()),model_ak72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[422] = 1;
  return result;
}

static ExcelValue _common423() {
  static ExcelValue result;
  if(variable_set[423] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_al75();
  array3[1] = model_al76();
  array3[2] = model_al77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_al75();
  array1[1] = model_al76();
  array1[2] = model_al77();
  array1[3] = subtract(model_al74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_al76(),model_al67()),subtract(C8,model_al72())))};
  ExcelValue array6[] = {model_al72(),subtract(model_al72(),multiply(divide(model_al77(),model_al68()),model_al72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_al74());
  variable_set[423] = 1;
  return result;
}

static ExcelValue _common424() {
  static ExcelValue result;
  if(variable_set[424] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_al75();
  array3[1] = model_al76();
  array3[2] = model_al77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_al75();
  array1[1] = model_al76();
  array1[2] = model_al77();
  array1[3] = subtract(model_al74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_al76(),model_al67()),subtract(C8,model_al72())))};
  ExcelValue array6[] = {model_al72(),subtract(model_al72(),multiply(divide(model_al77(),model_al68()),model_al72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[424] = 1;
  return result;
}

static ExcelValue _common425() {
  static ExcelValue result;
  if(variable_set[425] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_al75();
  array2[1] = model_al76();
  array2[2] = model_al77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_al75();
  array0[1] = model_al76();
  array0[2] = model_al77();
  array0[3] = subtract(model_al74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[425] = 1;
  return result;
}

static ExcelValue _common426() {
  static ExcelValue result;
  if(variable_set[426] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_al76(),model_al67()),subtract(C8,model_al72())))};
  ExcelValue array2[] = {model_al72(),subtract(model_al72(),multiply(divide(model_al77(),model_al68()),model_al72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[426] = 1;
  return result;
}

static ExcelValue _common427() {
  static ExcelValue result;
  if(variable_set[427] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_am75();
  array3[1] = model_am76();
  array3[2] = model_am77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_am75();
  array1[1] = model_am76();
  array1[2] = model_am77();
  array1[3] = subtract(model_am74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_am76(),model_am67()),subtract(C8,model_am72())))};
  ExcelValue array6[] = {model_am72(),subtract(model_am72(),multiply(divide(model_am77(),model_am68()),model_am72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_am74());
  variable_set[427] = 1;
  return result;
}

static ExcelValue _common428() {
  static ExcelValue result;
  if(variable_set[428] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_am75();
  array3[1] = model_am76();
  array3[2] = model_am77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_am75();
  array1[1] = model_am76();
  array1[2] = model_am77();
  array1[3] = subtract(model_am74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_am76(),model_am67()),subtract(C8,model_am72())))};
  ExcelValue array6[] = {model_am72(),subtract(model_am72(),multiply(divide(model_am77(),model_am68()),model_am72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[428] = 1;
  return result;
}

static ExcelValue _common429() {
  static ExcelValue result;
  if(variable_set[429] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_am75();
  array2[1] = model_am76();
  array2[2] = model_am77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_am75();
  array0[1] = model_am76();
  array0[2] = model_am77();
  array0[3] = subtract(model_am74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[429] = 1;
  return result;
}

static ExcelValue _common430() {
  static ExcelValue result;
  if(variable_set[430] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_am76(),model_am67()),subtract(C8,model_am72())))};
  ExcelValue array2[] = {model_am72(),subtract(model_am72(),multiply(divide(model_am77(),model_am68()),model_am72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[430] = 1;
  return result;
}

static ExcelValue _common431() {
  static ExcelValue result;
  if(variable_set[431] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_an75();
  array3[1] = model_an76();
  array3[2] = model_an77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_an75();
  array1[1] = model_an76();
  array1[2] = model_an77();
  array1[3] = subtract(model_an74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_an76(),model_an67()),subtract(C8,model_an72())))};
  ExcelValue array6[] = {model_an72(),subtract(model_an72(),multiply(divide(model_an77(),model_an68()),model_an72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = divide(sumproduct(2, array0),model_an74());
  variable_set[431] = 1;
  return result;
}

static ExcelValue _common432() {
  static ExcelValue result;
  if(variable_set[432] == 1) { return result;}
  static ExcelValue array3[3];
  array3[0] = model_an75();
  array3[1] = model_an76();
  array3[2] = model_an77();
  ExcelValue array3_ev = new_excel_range(array3,3,1);
  ExcelValue array2[] = {array3_ev};
  static ExcelValue array1[4];
  array1[0] = model_an75();
  array1[1] = model_an76();
  array1[2] = model_an77();
  array1[3] = subtract(model_an74(),sum(1, array2));
  ExcelValue array1_ev = new_excel_range(array1,4,1);
  ExcelValue array5[] = {C8,subtract(C8,multiply(divide(model_an76(),model_an67()),subtract(C8,model_an72())))};
  ExcelValue array6[] = {model_an72(),subtract(model_an72(),multiply(divide(model_an77(),model_an68()),model_an72()))};
  static ExcelValue array4[4];
  array4[0] = C8;
  array4[1] = average(2, array5);
  array4[2] = average(2, array6);
  array4[3] = C37;
  ExcelValue array4_ev = new_excel_range(array4,4,1);
  ExcelValue array0[] = {array1_ev,array4_ev};
  result = sumproduct(2, array0);
  variable_set[432] = 1;
  return result;
}

static ExcelValue _common433() {
  static ExcelValue result;
  if(variable_set[433] == 1) { return result;}
  static ExcelValue array2[3];
  array2[0] = model_an75();
  array2[1] = model_an76();
  array2[2] = model_an77();
  ExcelValue array2_ev = new_excel_range(array2,3,1);
  ExcelValue array1[] = {array2_ev};
  static ExcelValue array0[4];
  array0[0] = model_an75();
  array0[1] = model_an76();
  array0[2] = model_an77();
  array0[3] = subtract(model_an74(),sum(1, array1));
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[433] = 1;
  return result;
}

static ExcelValue _common434() {
  static ExcelValue result;
  if(variable_set[434] == 1) { return result;}
  ExcelValue array1[] = {C8,subtract(C8,multiply(divide(model_an76(),model_an67()),subtract(C8,model_an72())))};
  ExcelValue array2[] = {model_an72(),subtract(model_an72(),multiply(divide(model_an77(),model_an68()),model_an72()))};
  static ExcelValue array0[4];
  array0[0] = C8;
  array0[1] = average(2, array1);
  array0[2] = average(2, array2);
  array0[3] = C37;
  ExcelValue array0_ev = new_excel_range(array0,4,1);
  result = array0_ev;
  variable_set[434] = 1;
  return result;
}

// ending common elements

// start Model
ExcelValue model_b3() {
  static ExcelValue result;
  if(variable_set[435] == 1) { return result;}
  result = C1;
  variable_set[435] = 1;
  return result;
}

ExcelValue model_f3() {
  static ExcelValue result;
  if(variable_set[436] == 1) { return result;}
  static ExcelValue array1[5];
  array1[0] = model_m53();
  array1[1] = model_n53();
  array1[2] = model_o53();
  array1[3] = model_p53();
  array1[4] = model_q53();
  ExcelValue array1_ev = new_excel_range(array1,1,5);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[436] = 1;
  return result;
}

ExcelValue model_b4() {
  static ExcelValue result;
  if(variable_set[437] == 1) { return result;}
  result = C2;
  variable_set[437] = 1;
  return result;
}

ExcelValue model_f6() {
  static ExcelValue result;
  if(variable_set[438] == 1) { return result;}
  result = _common0();
  variable_set[438] = 1;
  return result;
}

ExcelValue model_b7() {
  static ExcelValue result;
  if(variable_set[439] == 1) { return result;}
  result = C3;
  variable_set[439] = 1;
  return result;
}

ExcelValue model_f7() {
  static ExcelValue result;
  if(variable_set[440] == 1) { return result;}
  result = _common5();
  variable_set[440] = 1;
  return result;
}

ExcelValue model_b8_default() {
  return C4;
}
static ExcelValue model_b8_variable;
ExcelValue model_b8() { if(variable_set[441] == 1) { return model_b8_variable; } else { return model_b8_default(); } }
void set_model_b8(ExcelValue newValue) { variable_set[441] = 1; model_b8_variable = newValue; }

ExcelValue model_b9_default() {
  return C5;
}
static ExcelValue model_b9_variable;
ExcelValue model_b9() { if(variable_set[442] == 1) { return model_b9_variable; } else { return model_b9_default(); } }
void set_model_b9(ExcelValue newValue) { variable_set[442] = 1; model_b9_variable = newValue; }

ExcelValue model_b10() {
  static ExcelValue result;
  if(variable_set[443] == 1) { return result;}
  result = C6;
  variable_set[443] = 1;
  return result;
}

ExcelValue model_b11() {
  static ExcelValue result;
  if(variable_set[444] == 1) { return result;}
  result = C7;
  variable_set[444] = 1;
  return result;
}

ExcelValue model_b12() {
  static ExcelValue result;
  if(variable_set[445] == 1) { return result;}
  result = C8;
  variable_set[445] = 1;
  return result;
}

ExcelValue model_b13() {
  static ExcelValue result;
  if(variable_set[446] == 1) { return result;}
  result = C9;
  variable_set[446] = 1;
  return result;
}

ExcelValue model_b31() {
  static ExcelValue result;
  if(variable_set[447] == 1) { return result;}
  result = C10;
  variable_set[447] = 1;
  return result;
}

ExcelValue model_b32() {
  static ExcelValue result;
  if(variable_set[448] == 1) { return result;}
  result = C11;
  variable_set[448] = 1;
  return result;
}

ExcelValue model_b34() {
  static ExcelValue result;
  if(variable_set[449] == 1) { return result;}
  result = C12;
  variable_set[449] = 1;
  return result;
}

ExcelValue model_b35() {
  static ExcelValue result;
  if(variable_set[450] == 1) { return result;}
  result = C13;
  variable_set[450] = 1;
  return result;
}

ExcelValue model_b36() {
  static ExcelValue result;
  if(variable_set[451] == 1) { return result;}
  result = C14;
  variable_set[451] = 1;
  return result;
}

ExcelValue model_b37() {
  static ExcelValue result;
  if(variable_set[452] == 1) { return result;}
  result = C15;
  variable_set[452] = 1;
  return result;
}

static ExcelValue model_n38() {
  static ExcelValue result;
  if(variable_set[453] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = multiply(C3,model_j48());
  array1[1] = C15;
  array1[2] = C16;
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = divide(subtract(sum(1, array0),C17),C18);
  variable_set[453] = 1;
  return result;
}

ExcelValue model_b40() {
  static ExcelValue result;
  if(variable_set[454] == 1) { return result;}
  result = C19;
  variable_set[454] = 1;
  return result;
}

ExcelValue model_c40() {
  static ExcelValue result;
  if(variable_set[455] == 1) { return result;}
  result = C20;
  variable_set[455] = 1;
  return result;
}

ExcelValue model_d40() {
  static ExcelValue result;
  if(variable_set[456] == 1) { return result;}
  result = C21;
  variable_set[456] = 1;
  return result;
}

ExcelValue model_b44() {
  static ExcelValue result;
  if(variable_set[457] == 1) { return result;}
  result = C7;
  variable_set[457] = 1;
  return result;
}

ExcelValue model_c44() {
  static ExcelValue result;
  if(variable_set[458] == 1) { return result;}
  result = C7;
  variable_set[458] = 1;
  return result;
}

ExcelValue model_b45() {
  static ExcelValue result;
  if(variable_set[459] == 1) { return result;}
  result = C22;
  variable_set[459] = 1;
  return result;
}

ExcelValue model_c45() {
  static ExcelValue result;
  if(variable_set[460] == 1) { return result;}
  result = C23;
  variable_set[460] = 1;
  return result;
}

static ExcelValue model_f47() {
  static ExcelValue result;
  if(variable_set[461] == 1) { return result;}
  result = add(C24,C8);
  variable_set[461] = 1;
  return result;
}

static ExcelValue model_g47() {
  static ExcelValue result;
  if(variable_set[462] == 1) { return result;}
  result = add(model_f47(),C8);
  variable_set[462] = 1;
  return result;
}

static ExcelValue model_h47() {
  static ExcelValue result;
  if(variable_set[463] == 1) { return result;}
  result = add(model_g47(),C8);
  variable_set[463] = 1;
  return result;
}

static ExcelValue model_i47() {
  static ExcelValue result;
  if(variable_set[464] == 1) { return result;}
  result = add(model_h47(),C8);
  variable_set[464] = 1;
  return result;
}

static ExcelValue model_j47() {
  static ExcelValue result;
  if(variable_set[465] == 1) { return result;}
  result = add(model_i47(),C8);
  variable_set[465] = 1;
  return result;
}

static ExcelValue model_k47() {
  static ExcelValue result;
  if(variable_set[466] == 1) { return result;}
  result = add(model_j47(),C8);
  variable_set[466] = 1;
  return result;
}

static ExcelValue model_l47() {
  static ExcelValue result;
  if(variable_set[467] == 1) { return result;}
  result = add(model_k47(),C8);
  variable_set[467] = 1;
  return result;
}

static ExcelValue model_m47() {
  static ExcelValue result;
  if(variable_set[468] == 1) { return result;}
  result = add(model_l47(),C8);
  variable_set[468] = 1;
  return result;
}

static ExcelValue model_n47() {
  static ExcelValue result;
  if(variable_set[469] == 1) { return result;}
  result = add(model_m47(),C8);
  variable_set[469] = 1;
  return result;
}

static ExcelValue model_o47() {
  static ExcelValue result;
  if(variable_set[470] == 1) { return result;}
  result = add(model_n47(),C8);
  variable_set[470] = 1;
  return result;
}

static ExcelValue model_p47() {
  static ExcelValue result;
  if(variable_set[471] == 1) { return result;}
  result = add(model_o47(),C8);
  variable_set[471] = 1;
  return result;
}

static ExcelValue model_q47() {
  static ExcelValue result;
  if(variable_set[472] == 1) { return result;}
  result = add(model_p47(),C8);
  variable_set[472] = 1;
  return result;
}

static ExcelValue model_r47() {
  static ExcelValue result;
  if(variable_set[473] == 1) { return result;}
  result = add(model_q47(),C8);
  variable_set[473] = 1;
  return result;
}

static ExcelValue model_s47() {
  static ExcelValue result;
  if(variable_set[474] == 1) { return result;}
  result = add(model_r47(),C8);
  variable_set[474] = 1;
  return result;
}

static ExcelValue model_t47() {
  static ExcelValue result;
  if(variable_set[475] == 1) { return result;}
  result = add(model_s47(),C8);
  variable_set[475] = 1;
  return result;
}

static ExcelValue model_u47() {
  static ExcelValue result;
  if(variable_set[476] == 1) { return result;}
  result = add(model_t47(),C8);
  variable_set[476] = 1;
  return result;
}

static ExcelValue model_v47() {
  static ExcelValue result;
  if(variable_set[477] == 1) { return result;}
  result = add(model_u47(),C8);
  variable_set[477] = 1;
  return result;
}

static ExcelValue model_w47() {
  static ExcelValue result;
  if(variable_set[478] == 1) { return result;}
  result = add(model_v47(),C8);
  variable_set[478] = 1;
  return result;
}

static ExcelValue model_x47() {
  static ExcelValue result;
  if(variable_set[479] == 1) { return result;}
  result = add(model_w47(),C8);
  variable_set[479] = 1;
  return result;
}

static ExcelValue model_y47() {
  static ExcelValue result;
  if(variable_set[480] == 1) { return result;}
  result = add(model_x47(),C8);
  variable_set[480] = 1;
  return result;
}

static ExcelValue model_z47() {
  static ExcelValue result;
  if(variable_set[481] == 1) { return result;}
  result = add(model_y47(),C8);
  variable_set[481] = 1;
  return result;
}

static ExcelValue model_aa47() {
  static ExcelValue result;
  if(variable_set[482] == 1) { return result;}
  result = add(model_z47(),C8);
  variable_set[482] = 1;
  return result;
}

static ExcelValue model_ab47() {
  static ExcelValue result;
  if(variable_set[483] == 1) { return result;}
  result = add(model_aa47(),C8);
  variable_set[483] = 1;
  return result;
}

static ExcelValue model_ac47() {
  static ExcelValue result;
  if(variable_set[484] == 1) { return result;}
  result = add(model_ab47(),C8);
  variable_set[484] = 1;
  return result;
}

static ExcelValue model_ad47() {
  static ExcelValue result;
  if(variable_set[485] == 1) { return result;}
  result = add(model_ac47(),C8);
  variable_set[485] = 1;
  return result;
}

static ExcelValue model_ae47() {
  static ExcelValue result;
  if(variable_set[486] == 1) { return result;}
  result = add(model_ad47(),C8);
  variable_set[486] = 1;
  return result;
}

static ExcelValue model_af47() {
  static ExcelValue result;
  if(variable_set[487] == 1) { return result;}
  result = add(model_ae47(),C8);
  variable_set[487] = 1;
  return result;
}

static ExcelValue model_ag47() {
  static ExcelValue result;
  if(variable_set[488] == 1) { return result;}
  result = add(model_af47(),C8);
  variable_set[488] = 1;
  return result;
}

static ExcelValue model_ah47() {
  static ExcelValue result;
  if(variable_set[489] == 1) { return result;}
  result = add(model_ag47(),C8);
  variable_set[489] = 1;
  return result;
}

static ExcelValue model_ai47() {
  static ExcelValue result;
  if(variable_set[490] == 1) { return result;}
  result = add(model_ah47(),C8);
  variable_set[490] = 1;
  return result;
}

static ExcelValue model_aj47() {
  static ExcelValue result;
  if(variable_set[491] == 1) { return result;}
  result = add(model_ai47(),C8);
  variable_set[491] = 1;
  return result;
}

static ExcelValue model_ak47() {
  static ExcelValue result;
  if(variable_set[492] == 1) { return result;}
  result = add(model_aj47(),C8);
  variable_set[492] = 1;
  return result;
}

static ExcelValue model_al47() {
  static ExcelValue result;
  if(variable_set[493] == 1) { return result;}
  result = add(model_ak47(),C8);
  variable_set[493] = 1;
  return result;
}

static ExcelValue model_am47() {
  static ExcelValue result;
  if(variable_set[494] == 1) { return result;}
  result = add(model_al47(),C8);
  variable_set[494] = 1;
  return result;
}

static ExcelValue model_an47() {
  static ExcelValue result;
  if(variable_set[495] == 1) { return result;}
  result = add(model_am47(),C8);
  variable_set[495] = 1;
  return result;
}

ExcelValue model_b48() {
  static ExcelValue result;
  if(variable_set[496] == 1) { return result;}
  result = C10;
  variable_set[496] = 1;
  return result;
}

ExcelValue model_c48() {
  static ExcelValue result;
  if(variable_set[497] == 1) { return result;}
  result = C25;
  variable_set[497] = 1;
  return result;
}

ExcelValue model_d48() {
  static ExcelValue result;
  if(variable_set[498] == 1) { return result;}
  result = multiply(C25,C26);
  variable_set[498] = 1;
  return result;
}

ExcelValue model_e48() {
  static ExcelValue result;
  if(variable_set[499] == 1) { return result;}
  result = multiply(model_d48(),excel_if(more_than(C24,C1),C27,C26));
  variable_set[499] = 1;
  return result;
}

ExcelValue model_f48() {
  static ExcelValue result;
  if(variable_set[500] == 1) { return result;}
  result = multiply(model_e48(),excel_if(more_than(model_f47(),C1),C27,C26));
  variable_set[500] = 1;
  return result;
}

ExcelValue model_g48() {
  static ExcelValue result;
  if(variable_set[501] == 1) { return result;}
  result = multiply(model_f48(),excel_if(more_than(model_g47(),C1),C27,C26));
  variable_set[501] = 1;
  return result;
}

ExcelValue model_h48() {
  static ExcelValue result;
  if(variable_set[502] == 1) { return result;}
  result = multiply(model_g48(),excel_if(more_than(model_h47(),C1),C27,C26));
  variable_set[502] = 1;
  return result;
}

ExcelValue model_i48() {
  static ExcelValue result;
  if(variable_set[503] == 1) { return result;}
  result = multiply(model_h48(),excel_if(more_than(model_i47(),C1),C27,C26));
  variable_set[503] = 1;
  return result;
}

ExcelValue model_j48() {
  static ExcelValue result;
  if(variable_set[504] == 1) { return result;}
  result = multiply(model_i48(),excel_if(more_than(model_j47(),C1),C27,C26));
  variable_set[504] = 1;
  return result;
}

ExcelValue model_k48() {
  static ExcelValue result;
  if(variable_set[505] == 1) { return result;}
  result = multiply(model_j48(),excel_if(more_than(model_k47(),C1),C27,C26));
  variable_set[505] = 1;
  return result;
}

ExcelValue model_l48() {
  static ExcelValue result;
  if(variable_set[506] == 1) { return result;}
  result = multiply(model_k48(),excel_if(more_than(model_l47(),C1),C27,C26));
  variable_set[506] = 1;
  return result;
}

ExcelValue model_m48() {
  static ExcelValue result;
  if(variable_set[507] == 1) { return result;}
  result = multiply(model_l48(),excel_if(more_than(model_m47(),C1),C27,C26));
  variable_set[507] = 1;
  return result;
}

ExcelValue model_n48() {
  static ExcelValue result;
  if(variable_set[508] == 1) { return result;}
  result = multiply(model_m48(),excel_if(more_than(model_n47(),C1),C27,C26));
  variable_set[508] = 1;
  return result;
}

ExcelValue model_o48() {
  static ExcelValue result;
  if(variable_set[509] == 1) { return result;}
  result = multiply(model_n48(),excel_if(more_than(model_o47(),C1),C27,C26));
  variable_set[509] = 1;
  return result;
}

ExcelValue model_p48() {
  static ExcelValue result;
  if(variable_set[510] == 1) { return result;}
  result = multiply(model_o48(),excel_if(more_than(model_p47(),C1),C27,C26));
  variable_set[510] = 1;
  return result;
}

ExcelValue model_q48() {
  static ExcelValue result;
  if(variable_set[511] == 1) { return result;}
  result = multiply(model_p48(),excel_if(more_than(model_q47(),C1),C27,C26));
  variable_set[511] = 1;
  return result;
}

ExcelValue model_r48() {
  static ExcelValue result;
  if(variable_set[512] == 1) { return result;}
  result = multiply(model_q48(),excel_if(more_than(model_r47(),C1),C27,C26));
  variable_set[512] = 1;
  return result;
}

ExcelValue model_s48() {
  static ExcelValue result;
  if(variable_set[513] == 1) { return result;}
  result = multiply(model_r48(),excel_if(more_than(model_s47(),C1),C27,C26));
  variable_set[513] = 1;
  return result;
}

ExcelValue model_t48() {
  static ExcelValue result;
  if(variable_set[514] == 1) { return result;}
  result = multiply(model_s48(),excel_if(more_than(model_t47(),C1),C27,C26));
  variable_set[514] = 1;
  return result;
}

ExcelValue model_u48() {
  static ExcelValue result;
  if(variable_set[515] == 1) { return result;}
  result = multiply(model_t48(),excel_if(more_than(model_u47(),C1),C27,C26));
  variable_set[515] = 1;
  return result;
}

ExcelValue model_v48() {
  static ExcelValue result;
  if(variable_set[516] == 1) { return result;}
  result = multiply(model_u48(),excel_if(more_than(model_v47(),C1),C27,C26));
  variable_set[516] = 1;
  return result;
}

ExcelValue model_w48() {
  static ExcelValue result;
  if(variable_set[517] == 1) { return result;}
  result = multiply(model_v48(),excel_if(more_than(model_w47(),C1),C27,C26));
  variable_set[517] = 1;
  return result;
}

ExcelValue model_x48() {
  static ExcelValue result;
  if(variable_set[518] == 1) { return result;}
  result = multiply(model_w48(),excel_if(more_than(model_x47(),C1),C27,C26));
  variable_set[518] = 1;
  return result;
}

ExcelValue model_y48() {
  static ExcelValue result;
  if(variable_set[519] == 1) { return result;}
  result = multiply(model_x48(),excel_if(more_than(model_y47(),C1),C27,C26));
  variable_set[519] = 1;
  return result;
}

ExcelValue model_z48() {
  static ExcelValue result;
  if(variable_set[520] == 1) { return result;}
  result = multiply(model_y48(),excel_if(more_than(model_z47(),C1),C27,C26));
  variable_set[520] = 1;
  return result;
}

ExcelValue model_aa48() {
  static ExcelValue result;
  if(variable_set[521] == 1) { return result;}
  result = multiply(model_z48(),excel_if(more_than(model_aa47(),C1),C27,C26));
  variable_set[521] = 1;
  return result;
}

ExcelValue model_ab48() {
  static ExcelValue result;
  if(variable_set[522] == 1) { return result;}
  result = multiply(model_aa48(),excel_if(more_than(model_ab47(),C1),C27,C26));
  variable_set[522] = 1;
  return result;
}

ExcelValue model_ac48() {
  static ExcelValue result;
  if(variable_set[523] == 1) { return result;}
  result = multiply(model_ab48(),excel_if(more_than(model_ac47(),C1),C27,C26));
  variable_set[523] = 1;
  return result;
}

ExcelValue model_ad48() {
  static ExcelValue result;
  if(variable_set[524] == 1) { return result;}
  result = multiply(model_ac48(),excel_if(more_than(model_ad47(),C1),C27,C26));
  variable_set[524] = 1;
  return result;
}

ExcelValue model_ae48() {
  static ExcelValue result;
  if(variable_set[525] == 1) { return result;}
  result = multiply(model_ad48(),excel_if(more_than(model_ae47(),C1),C27,C26));
  variable_set[525] = 1;
  return result;
}

ExcelValue model_af48() {
  static ExcelValue result;
  if(variable_set[526] == 1) { return result;}
  result = multiply(model_ae48(),excel_if(more_than(model_af47(),C1),C27,C26));
  variable_set[526] = 1;
  return result;
}

ExcelValue model_ag48() {
  static ExcelValue result;
  if(variable_set[527] == 1) { return result;}
  result = multiply(model_af48(),excel_if(more_than(model_ag47(),C1),C27,C26));
  variable_set[527] = 1;
  return result;
}

ExcelValue model_ah48() {
  static ExcelValue result;
  if(variable_set[528] == 1) { return result;}
  result = multiply(model_ag48(),excel_if(more_than(model_ah47(),C1),C27,C26));
  variable_set[528] = 1;
  return result;
}

ExcelValue model_ai48() {
  static ExcelValue result;
  if(variable_set[529] == 1) { return result;}
  result = multiply(model_ah48(),excel_if(more_than(model_ai47(),C1),C27,C26));
  variable_set[529] = 1;
  return result;
}

ExcelValue model_aj48() {
  static ExcelValue result;
  if(variable_set[530] == 1) { return result;}
  result = multiply(model_ai48(),excel_if(more_than(model_aj47(),C1),C27,C26));
  variable_set[530] = 1;
  return result;
}

ExcelValue model_ak48() {
  static ExcelValue result;
  if(variable_set[531] == 1) { return result;}
  result = multiply(model_aj48(),excel_if(more_than(model_ak47(),C1),C27,C26));
  variable_set[531] = 1;
  return result;
}

ExcelValue model_al48() {
  static ExcelValue result;
  if(variable_set[532] == 1) { return result;}
  result = multiply(model_ak48(),excel_if(more_than(model_al47(),C1),C27,C26));
  variable_set[532] = 1;
  return result;
}

ExcelValue model_am48() {
  static ExcelValue result;
  if(variable_set[533] == 1) { return result;}
  result = multiply(model_al48(),excel_if(more_than(model_am47(),C1),C27,C26));
  variable_set[533] = 1;
  return result;
}

ExcelValue model_an48() {
  static ExcelValue result;
  if(variable_set[534] == 1) { return result;}
  result = multiply(model_am48(),excel_if(more_than(model_an47(),C1),C27,C26));
  variable_set[534] = 1;
  return result;
}

ExcelValue model_b49() {
  static ExcelValue result;
  if(variable_set[535] == 1) { return result;}
  result = C17;
  variable_set[535] = 1;
  return result;
}

ExcelValue model_c49() {
  static ExcelValue result;
  if(variable_set[536] == 1) { return result;}
  result = add(C17,model_n38());
  variable_set[536] = 1;
  return result;
}

ExcelValue model_d49() {
  static ExcelValue result;
  if(variable_set[537] == 1) { return result;}
  result = add(model_c49(),model_n38());
  variable_set[537] = 1;
  return result;
}

ExcelValue model_e49() {
  static ExcelValue result;
  if(variable_set[538] == 1) { return result;}
  result = add(model_d49(),model_n38());
  variable_set[538] = 1;
  return result;
}

ExcelValue model_f49() {
  static ExcelValue result;
  if(variable_set[539] == 1) { return result;}
  result = add(model_e49(),model_n38());
  variable_set[539] = 1;
  return result;
}

ExcelValue model_g49() {
  static ExcelValue result;
  if(variable_set[540] == 1) { return result;}
  result = add(model_f49(),model_n38());
  variable_set[540] = 1;
  return result;
}

ExcelValue model_h49() {
  static ExcelValue result;
  if(variable_set[541] == 1) { return result;}
  result = add(model_g49(),model_n38());
  variable_set[541] = 1;
  return result;
}

ExcelValue model_i49() {
  static ExcelValue result;
  if(variable_set[542] == 1) { return result;}
  result = add(model_h49(),model_n38());
  variable_set[542] = 1;
  return result;
}

ExcelValue model_j49() {
  static ExcelValue result;
  if(variable_set[543] == 1) { return result;}
  result = add(model_i49(),model_n38());
  variable_set[543] = 1;
  return result;
}

ExcelValue model_k49() {
  static ExcelValue result;
  if(variable_set[544] == 1) { return result;}
  result = model_k86();
  variable_set[544] = 1;
  return result;
}

ExcelValue model_l49() {
  static ExcelValue result;
  if(variable_set[545] == 1) { return result;}
  result = model_l86();
  variable_set[545] = 1;
  return result;
}

ExcelValue model_m49() {
  static ExcelValue result;
  if(variable_set[546] == 1) { return result;}
  result = model_m86();
  variable_set[546] = 1;
  return result;
}

ExcelValue model_n49() {
  static ExcelValue result;
  if(variable_set[547] == 1) { return result;}
  result = model_n86();
  variable_set[547] = 1;
  return result;
}

ExcelValue model_o49() {
  static ExcelValue result;
  if(variable_set[548] == 1) { return result;}
  result = model_o86();
  variable_set[548] = 1;
  return result;
}

ExcelValue model_p49() {
  static ExcelValue result;
  if(variable_set[549] == 1) { return result;}
  result = model_p86();
  variable_set[549] = 1;
  return result;
}

ExcelValue model_q49() {
  static ExcelValue result;
  if(variable_set[550] == 1) { return result;}
  result = model_q86();
  variable_set[550] = 1;
  return result;
}

ExcelValue model_r49() {
  static ExcelValue result;
  if(variable_set[551] == 1) { return result;}
  result = model_r86();
  variable_set[551] = 1;
  return result;
}

ExcelValue model_s49() {
  static ExcelValue result;
  if(variable_set[552] == 1) { return result;}
  result = model_s86();
  variable_set[552] = 1;
  return result;
}

ExcelValue model_t49() {
  static ExcelValue result;
  if(variable_set[553] == 1) { return result;}
  result = model_t86();
  variable_set[553] = 1;
  return result;
}

ExcelValue model_u49() {
  static ExcelValue result;
  if(variable_set[554] == 1) { return result;}
  result = model_u86();
  variable_set[554] = 1;
  return result;
}

ExcelValue model_v49() {
  static ExcelValue result;
  if(variable_set[555] == 1) { return result;}
  result = model_v86();
  variable_set[555] = 1;
  return result;
}

ExcelValue model_w49() {
  static ExcelValue result;
  if(variable_set[556] == 1) { return result;}
  result = model_w86();
  variable_set[556] = 1;
  return result;
}

ExcelValue model_x49() {
  static ExcelValue result;
  if(variable_set[557] == 1) { return result;}
  result = model_x86();
  variable_set[557] = 1;
  return result;
}

ExcelValue model_y49() {
  static ExcelValue result;
  if(variable_set[558] == 1) { return result;}
  result = model_y86();
  variable_set[558] = 1;
  return result;
}

ExcelValue model_z49() {
  static ExcelValue result;
  if(variable_set[559] == 1) { return result;}
  result = model_z86();
  variable_set[559] = 1;
  return result;
}

ExcelValue model_aa49() {
  static ExcelValue result;
  if(variable_set[560] == 1) { return result;}
  result = model_aa86();
  variable_set[560] = 1;
  return result;
}

ExcelValue model_ab49() {
  static ExcelValue result;
  if(variable_set[561] == 1) { return result;}
  result = model_ab86();
  variable_set[561] = 1;
  return result;
}

ExcelValue model_ac49() {
  static ExcelValue result;
  if(variable_set[562] == 1) { return result;}
  result = model_ac86();
  variable_set[562] = 1;
  return result;
}

ExcelValue model_ad49() {
  static ExcelValue result;
  if(variable_set[563] == 1) { return result;}
  result = model_ad86();
  variable_set[563] = 1;
  return result;
}

ExcelValue model_ae49() {
  static ExcelValue result;
  if(variable_set[564] == 1) { return result;}
  result = model_ae86();
  variable_set[564] = 1;
  return result;
}

ExcelValue model_af49() {
  static ExcelValue result;
  if(variable_set[565] == 1) { return result;}
  result = model_af86();
  variable_set[565] = 1;
  return result;
}

ExcelValue model_ag49() {
  static ExcelValue result;
  if(variable_set[566] == 1) { return result;}
  result = model_ag86();
  variable_set[566] = 1;
  return result;
}

ExcelValue model_ah49() {
  static ExcelValue result;
  if(variable_set[567] == 1) { return result;}
  result = model_ah86();
  variable_set[567] = 1;
  return result;
}

ExcelValue model_ai49() {
  static ExcelValue result;
  if(variable_set[568] == 1) { return result;}
  result = model_ai86();
  variable_set[568] = 1;
  return result;
}

ExcelValue model_aj49() {
  static ExcelValue result;
  if(variable_set[569] == 1) { return result;}
  result = model_aj86();
  variable_set[569] = 1;
  return result;
}

ExcelValue model_ak49() {
  static ExcelValue result;
  if(variable_set[570] == 1) { return result;}
  result = model_ak86();
  variable_set[570] = 1;
  return result;
}

ExcelValue model_al49() {
  static ExcelValue result;
  if(variable_set[571] == 1) { return result;}
  result = model_al86();
  variable_set[571] = 1;
  return result;
}

ExcelValue model_am49() {
  static ExcelValue result;
  if(variable_set[572] == 1) { return result;}
  result = model_am86();
  variable_set[572] = 1;
  return result;
}

ExcelValue model_an49() {
  static ExcelValue result;
  if(variable_set[573] == 1) { return result;}
  result = model_an86();
  variable_set[573] = 1;
  return result;
}

ExcelValue model_b50() {
  static ExcelValue result;
  if(variable_set[574] == 1) { return result;}
  result = C28;
  variable_set[574] = 1;
  return result;
}

ExcelValue model_c50() {
  static ExcelValue result;
  if(variable_set[575] == 1) { return result;}
  result = _common11();
  variable_set[575] = 1;
  return result;
}

ExcelValue model_d50() {
  static ExcelValue result;
  if(variable_set[576] == 1) { return result;}
  result = _common12();
  variable_set[576] = 1;
  return result;
}

ExcelValue model_e50() {
  static ExcelValue result;
  if(variable_set[577] == 1) { return result;}
  result = _common13();
  variable_set[577] = 1;
  return result;
}

ExcelValue model_f50() {
  static ExcelValue result;
  if(variable_set[578] == 1) { return result;}
  result = _common14();
  variable_set[578] = 1;
  return result;
}

ExcelValue model_g50() {
  static ExcelValue result;
  if(variable_set[579] == 1) { return result;}
  result = _common15();
  variable_set[579] = 1;
  return result;
}

ExcelValue model_h50() {
  static ExcelValue result;
  if(variable_set[580] == 1) { return result;}
  result = _common16();
  variable_set[580] = 1;
  return result;
}

ExcelValue model_i50() {
  static ExcelValue result;
  if(variable_set[581] == 1) { return result;}
  result = _common17();
  variable_set[581] = 1;
  return result;
}

ExcelValue model_j50() {
  static ExcelValue result;
  if(variable_set[582] == 1) { return result;}
  result = _common18();
  variable_set[582] = 1;
  return result;
}

ExcelValue model_k50() {
  static ExcelValue result;
  if(variable_set[583] == 1) { return result;}
  result = _common19();
  variable_set[583] = 1;
  return result;
}

ExcelValue model_l50() {
  static ExcelValue result;
  if(variable_set[584] == 1) { return result;}
  result = _common20();
  variable_set[584] = 1;
  return result;
}

ExcelValue model_m50() {
  static ExcelValue result;
  if(variable_set[585] == 1) { return result;}
  result = _common21();
  variable_set[585] = 1;
  return result;
}

ExcelValue model_n50() {
  static ExcelValue result;
  if(variable_set[586] == 1) { return result;}
  result = _common22();
  variable_set[586] = 1;
  return result;
}

ExcelValue model_o50() {
  static ExcelValue result;
  if(variable_set[587] == 1) { return result;}
  result = _common23();
  variable_set[587] = 1;
  return result;
}

ExcelValue model_p50() {
  static ExcelValue result;
  if(variable_set[588] == 1) { return result;}
  result = _common24();
  variable_set[588] = 1;
  return result;
}

ExcelValue model_q50() {
  static ExcelValue result;
  if(variable_set[589] == 1) { return result;}
  result = _common25();
  variable_set[589] = 1;
  return result;
}

ExcelValue model_r50() {
  static ExcelValue result;
  if(variable_set[590] == 1) { return result;}
  result = _common26();
  variable_set[590] = 1;
  return result;
}

ExcelValue model_s50() {
  static ExcelValue result;
  if(variable_set[591] == 1) { return result;}
  result = _common27();
  variable_set[591] = 1;
  return result;
}

ExcelValue model_t50() {
  static ExcelValue result;
  if(variable_set[592] == 1) { return result;}
  result = _common4();
  variable_set[592] = 1;
  return result;
}

ExcelValue model_u50() {
  static ExcelValue result;
  if(variable_set[593] == 1) { return result;}
  result = _common28();
  variable_set[593] = 1;
  return result;
}

ExcelValue model_v50() {
  static ExcelValue result;
  if(variable_set[594] == 1) { return result;}
  result = _common29();
  variable_set[594] = 1;
  return result;
}

ExcelValue model_w50() {
  static ExcelValue result;
  if(variable_set[595] == 1) { return result;}
  result = _common30();
  variable_set[595] = 1;
  return result;
}

ExcelValue model_x50() {
  static ExcelValue result;
  if(variable_set[596] == 1) { return result;}
  result = _common31();
  variable_set[596] = 1;
  return result;
}

ExcelValue model_y50() {
  static ExcelValue result;
  if(variable_set[597] == 1) { return result;}
  result = _common32();
  variable_set[597] = 1;
  return result;
}

ExcelValue model_z50() {
  static ExcelValue result;
  if(variable_set[598] == 1) { return result;}
  result = _common33();
  variable_set[598] = 1;
  return result;
}

ExcelValue model_aa50() {
  static ExcelValue result;
  if(variable_set[599] == 1) { return result;}
  result = _common34();
  variable_set[599] = 1;
  return result;
}

ExcelValue model_ab50() {
  static ExcelValue result;
  if(variable_set[600] == 1) { return result;}
  result = _common35();
  variable_set[600] = 1;
  return result;
}

ExcelValue model_ac50() {
  static ExcelValue result;
  if(variable_set[601] == 1) { return result;}
  result = _common36();
  variable_set[601] = 1;
  return result;
}

ExcelValue model_ad50() {
  static ExcelValue result;
  if(variable_set[602] == 1) { return result;}
  result = _common37();
  variable_set[602] = 1;
  return result;
}

ExcelValue model_ae50() {
  static ExcelValue result;
  if(variable_set[603] == 1) { return result;}
  result = _common38();
  variable_set[603] = 1;
  return result;
}

ExcelValue model_af50() {
  static ExcelValue result;
  if(variable_set[604] == 1) { return result;}
  result = _common39();
  variable_set[604] = 1;
  return result;
}

ExcelValue model_ag50() {
  static ExcelValue result;
  if(variable_set[605] == 1) { return result;}
  result = _common40();
  variable_set[605] = 1;
  return result;
}

ExcelValue model_ah50() {
  static ExcelValue result;
  if(variable_set[606] == 1) { return result;}
  result = _common41();
  variable_set[606] = 1;
  return result;
}

ExcelValue model_ai50() {
  static ExcelValue result;
  if(variable_set[607] == 1) { return result;}
  result = _common42();
  variable_set[607] = 1;
  return result;
}

ExcelValue model_aj50() {
  static ExcelValue result;
  if(variable_set[608] == 1) { return result;}
  result = _common43();
  variable_set[608] = 1;
  return result;
}

ExcelValue model_ak50() {
  static ExcelValue result;
  if(variable_set[609] == 1) { return result;}
  result = _common44();
  variable_set[609] = 1;
  return result;
}

ExcelValue model_al50() {
  static ExcelValue result;
  if(variable_set[610] == 1) { return result;}
  result = _common45();
  variable_set[610] = 1;
  return result;
}

ExcelValue model_am50() {
  static ExcelValue result;
  if(variable_set[611] == 1) { return result;}
  result = _common46();
  variable_set[611] = 1;
  return result;
}

ExcelValue model_an50() {
  static ExcelValue result;
  if(variable_set[612] == 1) { return result;}
  result = _common10();
  variable_set[612] = 1;
  return result;
}

ExcelValue model_b51() {
  static ExcelValue result;
  if(variable_set[613] == 1) { return result;}
  result = C19;
  variable_set[613] = 1;
  return result;
}

ExcelValue model_c51() {
  static ExcelValue result;
  if(variable_set[614] == 1) { return result;}
  result = C29;
  variable_set[614] = 1;
  return result;
}

ExcelValue model_d51() {
  static ExcelValue result;
  if(variable_set[615] == 1) { return result;}
  result = C30;
  variable_set[615] = 1;
  return result;
}

ExcelValue model_e51() {
  static ExcelValue result;
  if(variable_set[616] == 1) { return result;}
  result = add(C30,C31);
  variable_set[616] = 1;
  return result;
}

ExcelValue model_f51() {
  static ExcelValue result;
  if(variable_set[617] == 1) { return result;}
  result = add(model_e51(),C31);
  variable_set[617] = 1;
  return result;
}

ExcelValue model_g51() {
  static ExcelValue result;
  if(variable_set[618] == 1) { return result;}
  result = add(model_f51(),C31);
  variable_set[618] = 1;
  return result;
}

ExcelValue model_h51() {
  static ExcelValue result;
  if(variable_set[619] == 1) { return result;}
  result = add(model_g51(),C31);
  variable_set[619] = 1;
  return result;
}

ExcelValue model_i51() {
  static ExcelValue result;
  if(variable_set[620] == 1) { return result;}
  result = add(model_h51(),C31);
  variable_set[620] = 1;
  return result;
}

ExcelValue model_j51() {
  static ExcelValue result;
  if(variable_set[621] == 1) { return result;}
  result = add(model_i51(),C31);
  variable_set[621] = 1;
  return result;
}

ExcelValue model_k51() {
  static ExcelValue result;
  if(variable_set[622] == 1) { return result;}
  result = add(model_j51(),C32);
  variable_set[622] = 1;
  return result;
}

ExcelValue model_l51() {
  static ExcelValue result;
  if(variable_set[623] == 1) { return result;}
  result = add(model_k51(),C32);
  variable_set[623] = 1;
  return result;
}

ExcelValue model_m51() {
  static ExcelValue result;
  if(variable_set[624] == 1) { return result;}
  result = add(model_l51(),C32);
  variable_set[624] = 1;
  return result;
}

ExcelValue model_n51() {
  static ExcelValue result;
  if(variable_set[625] == 1) { return result;}
  result = add(model_m51(),C32);
  variable_set[625] = 1;
  return result;
}

ExcelValue model_o51() {
  static ExcelValue result;
  if(variable_set[626] == 1) { return result;}
  result = add(model_n51(),C32);
  variable_set[626] = 1;
  return result;
}

ExcelValue model_p51() {
  static ExcelValue result;
  if(variable_set[627] == 1) { return result;}
  result = add(model_o51(),C32);
  variable_set[627] = 1;
  return result;
}

ExcelValue model_q51() {
  static ExcelValue result;
  if(variable_set[628] == 1) { return result;}
  result = add(model_p51(),C32);
  variable_set[628] = 1;
  return result;
}

ExcelValue model_r51() {
  static ExcelValue result;
  if(variable_set[629] == 1) { return result;}
  result = add(model_q51(),C32);
  variable_set[629] = 1;
  return result;
}

ExcelValue model_s51() {
  static ExcelValue result;
  if(variable_set[630] == 1) { return result;}
  result = add(model_r51(),C32);
  variable_set[630] = 1;
  return result;
}

ExcelValue model_t51() {
  static ExcelValue result;
  if(variable_set[631] == 1) { return result;}
  result = add(model_s51(),C32);
  variable_set[631] = 1;
  return result;
}

ExcelValue model_u51() {
  static ExcelValue result;
  if(variable_set[632] == 1) { return result;}
  result = add(model_t51(),C32);
  variable_set[632] = 1;
  return result;
}

ExcelValue model_v51() {
  static ExcelValue result;
  if(variable_set[633] == 1) { return result;}
  result = add(model_u51(),C32);
  variable_set[633] = 1;
  return result;
}

ExcelValue model_w51() {
  static ExcelValue result;
  if(variable_set[634] == 1) { return result;}
  result = add(model_v51(),C32);
  variable_set[634] = 1;
  return result;
}

ExcelValue model_x51() {
  static ExcelValue result;
  if(variable_set[635] == 1) { return result;}
  result = add(model_w51(),C32);
  variable_set[635] = 1;
  return result;
}

ExcelValue model_y51() {
  static ExcelValue result;
  if(variable_set[636] == 1) { return result;}
  result = add(model_x51(),C32);
  variable_set[636] = 1;
  return result;
}

ExcelValue model_z51() {
  static ExcelValue result;
  if(variable_set[637] == 1) { return result;}
  result = add(model_y51(),C32);
  variable_set[637] = 1;
  return result;
}

ExcelValue model_aa51() {
  static ExcelValue result;
  if(variable_set[638] == 1) { return result;}
  result = add(model_z51(),C32);
  variable_set[638] = 1;
  return result;
}

ExcelValue model_ab51() {
  static ExcelValue result;
  if(variable_set[639] == 1) { return result;}
  result = add(model_aa51(),C32);
  variable_set[639] = 1;
  return result;
}

ExcelValue model_ac51() {
  static ExcelValue result;
  if(variable_set[640] == 1) { return result;}
  result = add(model_ab51(),C32);
  variable_set[640] = 1;
  return result;
}

ExcelValue model_ad51() {
  static ExcelValue result;
  if(variable_set[641] == 1) { return result;}
  result = add(model_ac51(),C32);
  variable_set[641] = 1;
  return result;
}

ExcelValue model_ae51() {
  static ExcelValue result;
  if(variable_set[642] == 1) { return result;}
  result = add(model_ad51(),C32);
  variable_set[642] = 1;
  return result;
}

ExcelValue model_af51() {
  static ExcelValue result;
  if(variable_set[643] == 1) { return result;}
  result = add(model_ae51(),C32);
  variable_set[643] = 1;
  return result;
}

ExcelValue model_ag51() {
  static ExcelValue result;
  if(variable_set[644] == 1) { return result;}
  result = add(model_af51(),C32);
  variable_set[644] = 1;
  return result;
}

ExcelValue model_ah51() {
  static ExcelValue result;
  if(variable_set[645] == 1) { return result;}
  result = add(model_ag51(),C32);
  variable_set[645] = 1;
  return result;
}

ExcelValue model_ai51() {
  static ExcelValue result;
  if(variable_set[646] == 1) { return result;}
  result = add(model_ah51(),C32);
  variable_set[646] = 1;
  return result;
}

ExcelValue model_aj51() {
  static ExcelValue result;
  if(variable_set[647] == 1) { return result;}
  result = add(model_ai51(),C32);
  variable_set[647] = 1;
  return result;
}

ExcelValue model_ak51() {
  static ExcelValue result;
  if(variable_set[648] == 1) { return result;}
  result = add(model_aj51(),C32);
  variable_set[648] = 1;
  return result;
}

ExcelValue model_al51() {
  static ExcelValue result;
  if(variable_set[649] == 1) { return result;}
  result = add(model_ak51(),C32);
  variable_set[649] = 1;
  return result;
}

ExcelValue model_am51() {
  static ExcelValue result;
  if(variable_set[650] == 1) { return result;}
  result = add(model_al51(),C32);
  variable_set[650] = 1;
  return result;
}

ExcelValue model_an51() {
  static ExcelValue result;
  if(variable_set[651] == 1) { return result;}
  result = _common9();
  variable_set[651] = 1;
  return result;
}

ExcelValue model_b52() {
  static ExcelValue result;
  if(variable_set[652] == 1) { return result;}
  result = multiply(divide(C33,C10),C34);
  variable_set[652] = 1;
  return result;
}

ExcelValue model_c52() {
  static ExcelValue result;
  if(variable_set[653] == 1) { return result;}
  result = multiply(divide(_common47(),C25),C34);
  variable_set[653] = 1;
  return result;
}

ExcelValue model_d52() {
  static ExcelValue result;
  if(variable_set[654] == 1) { return result;}
  result = multiply(divide(_common49(),model_d48()),C34);
  variable_set[654] = 1;
  return result;
}

ExcelValue model_e52() {
  static ExcelValue result;
  if(variable_set[655] == 1) { return result;}
  result = multiply(divide(_common51(),model_e48()),C34);
  variable_set[655] = 1;
  return result;
}

ExcelValue model_f52() {
  static ExcelValue result;
  if(variable_set[656] == 1) { return result;}
  result = multiply(divide(_common53(),model_f48()),C34);
  variable_set[656] = 1;
  return result;
}

ExcelValue model_g52() {
  static ExcelValue result;
  if(variable_set[657] == 1) { return result;}
  result = multiply(divide(_common55(),model_g48()),C34);
  variable_set[657] = 1;
  return result;
}

ExcelValue model_h52() {
  static ExcelValue result;
  if(variable_set[658] == 1) { return result;}
  result = multiply(divide(_common57(),model_h48()),C34);
  variable_set[658] = 1;
  return result;
}

ExcelValue model_i52() {
  static ExcelValue result;
  if(variable_set[659] == 1) { return result;}
  result = multiply(divide(_common59(),model_i48()),C34);
  variable_set[659] = 1;
  return result;
}

ExcelValue model_j52() {
  static ExcelValue result;
  if(variable_set[660] == 1) { return result;}
  result = multiply(divide(_common61(),model_j48()),C34);
  variable_set[660] = 1;
  return result;
}

ExcelValue model_k52() {
  static ExcelValue result;
  if(variable_set[661] == 1) { return result;}
  result = multiply(divide(_common63(),model_k48()),C34);
  variable_set[661] = 1;
  return result;
}

ExcelValue model_l52() {
  static ExcelValue result;
  if(variable_set[662] == 1) { return result;}
  result = multiply(divide(_common65(),model_l48()),C34);
  variable_set[662] = 1;
  return result;
}

ExcelValue model_m52() {
  static ExcelValue result;
  if(variable_set[663] == 1) { return result;}
  result = multiply(divide(model_m53(),model_m48()),C34);
  variable_set[663] = 1;
  return result;
}

ExcelValue model_n52() {
  static ExcelValue result;
  if(variable_set[664] == 1) { return result;}
  result = multiply(divide(model_n53(),model_n48()),C34);
  variable_set[664] = 1;
  return result;
}

ExcelValue model_o52() {
  static ExcelValue result;
  if(variable_set[665] == 1) { return result;}
  result = multiply(divide(model_o53(),model_o48()),C34);
  variable_set[665] = 1;
  return result;
}

ExcelValue model_p52() {
  static ExcelValue result;
  if(variable_set[666] == 1) { return result;}
  result = multiply(divide(model_p53(),model_p48()),C34);
  variable_set[666] = 1;
  return result;
}

ExcelValue model_q52() {
  static ExcelValue result;
  if(variable_set[667] == 1) { return result;}
  result = multiply(divide(model_q53(),model_q48()),C34);
  variable_set[667] = 1;
  return result;
}

ExcelValue model_r52() {
  static ExcelValue result;
  if(variable_set[668] == 1) { return result;}
  result = multiply(divide(_common67(),model_r48()),C34);
  variable_set[668] = 1;
  return result;
}

ExcelValue model_s52() {
  static ExcelValue result;
  if(variable_set[669] == 1) { return result;}
  result = multiply(divide(_common69(),model_s48()),C34);
  variable_set[669] = 1;
  return result;
}

ExcelValue model_t52() {
  static ExcelValue result;
  if(variable_set[670] == 1) { return result;}
  result = _common0();
  variable_set[670] = 1;
  return result;
}

ExcelValue model_u52() {
  static ExcelValue result;
  if(variable_set[671] == 1) { return result;}
  result = multiply(divide(_common71(),model_u48()),C34);
  variable_set[671] = 1;
  return result;
}

ExcelValue model_v52() {
  static ExcelValue result;
  if(variable_set[672] == 1) { return result;}
  result = multiply(divide(_common73(),model_v48()),C34);
  variable_set[672] = 1;
  return result;
}

ExcelValue model_w52() {
  static ExcelValue result;
  if(variable_set[673] == 1) { return result;}
  result = multiply(divide(_common75(),model_w48()),C34);
  variable_set[673] = 1;
  return result;
}

ExcelValue model_x52() {
  static ExcelValue result;
  if(variable_set[674] == 1) { return result;}
  result = multiply(divide(_common77(),model_x48()),C34);
  variable_set[674] = 1;
  return result;
}

ExcelValue model_y52() {
  static ExcelValue result;
  if(variable_set[675] == 1) { return result;}
  result = multiply(divide(_common79(),model_y48()),C34);
  variable_set[675] = 1;
  return result;
}

ExcelValue model_z52() {
  static ExcelValue result;
  if(variable_set[676] == 1) { return result;}
  result = multiply(divide(_common81(),model_z48()),C34);
  variable_set[676] = 1;
  return result;
}

ExcelValue model_aa52() {
  static ExcelValue result;
  if(variable_set[677] == 1) { return result;}
  result = multiply(divide(_common83(),model_aa48()),C34);
  variable_set[677] = 1;
  return result;
}

ExcelValue model_ab52() {
  static ExcelValue result;
  if(variable_set[678] == 1) { return result;}
  result = multiply(divide(_common85(),model_ab48()),C34);
  variable_set[678] = 1;
  return result;
}

ExcelValue model_ac52() {
  static ExcelValue result;
  if(variable_set[679] == 1) { return result;}
  result = multiply(divide(_common87(),model_ac48()),C34);
  variable_set[679] = 1;
  return result;
}

ExcelValue model_ad52() {
  static ExcelValue result;
  if(variable_set[680] == 1) { return result;}
  result = multiply(divide(_common89(),model_ad48()),C34);
  variable_set[680] = 1;
  return result;
}

ExcelValue model_ae52() {
  static ExcelValue result;
  if(variable_set[681] == 1) { return result;}
  result = multiply(divide(_common91(),model_ae48()),C34);
  variable_set[681] = 1;
  return result;
}

ExcelValue model_af52() {
  static ExcelValue result;
  if(variable_set[682] == 1) { return result;}
  result = multiply(divide(_common93(),model_af48()),C34);
  variable_set[682] = 1;
  return result;
}

ExcelValue model_ag52() {
  static ExcelValue result;
  if(variable_set[683] == 1) { return result;}
  result = multiply(divide(_common95(),model_ag48()),C34);
  variable_set[683] = 1;
  return result;
}

ExcelValue model_ah52() {
  static ExcelValue result;
  if(variable_set[684] == 1) { return result;}
  result = multiply(divide(_common97(),model_ah48()),C34);
  variable_set[684] = 1;
  return result;
}

ExcelValue model_ai52() {
  static ExcelValue result;
  if(variable_set[685] == 1) { return result;}
  result = multiply(divide(_common99(),model_ai48()),C34);
  variable_set[685] = 1;
  return result;
}

ExcelValue model_aj52() {
  static ExcelValue result;
  if(variable_set[686] == 1) { return result;}
  result = multiply(divide(_common101(),model_aj48()),C34);
  variable_set[686] = 1;
  return result;
}

ExcelValue model_ak52() {
  static ExcelValue result;
  if(variable_set[687] == 1) { return result;}
  result = multiply(divide(_common103(),model_ak48()),C34);
  variable_set[687] = 1;
  return result;
}

ExcelValue model_al52() {
  static ExcelValue result;
  if(variable_set[688] == 1) { return result;}
  result = multiply(divide(_common105(),model_al48()),C34);
  variable_set[688] = 1;
  return result;
}

ExcelValue model_am52() {
  static ExcelValue result;
  if(variable_set[689] == 1) { return result;}
  result = multiply(divide(_common107(),model_am48()),C34);
  variable_set[689] = 1;
  return result;
}

ExcelValue model_an52() {
  static ExcelValue result;
  if(variable_set[690] == 1) { return result;}
  result = _common5();
  variable_set[690] = 1;
  return result;
}

ExcelValue model_b53() {
  static ExcelValue result;
  if(variable_set[691] == 1) { return result;}
  result = C33;
  variable_set[691] = 1;
  return result;
}

ExcelValue model_c53() {
  static ExcelValue result;
  if(variable_set[692] == 1) { return result;}
  result = _common47();
  variable_set[692] = 1;
  return result;
}

ExcelValue model_d53() {
  static ExcelValue result;
  if(variable_set[693] == 1) { return result;}
  result = _common49();
  variable_set[693] = 1;
  return result;
}

ExcelValue model_e53() {
  static ExcelValue result;
  if(variable_set[694] == 1) { return result;}
  result = _common51();
  variable_set[694] = 1;
  return result;
}

ExcelValue model_f53() {
  static ExcelValue result;
  if(variable_set[695] == 1) { return result;}
  result = _common53();
  variable_set[695] = 1;
  return result;
}

ExcelValue model_g53() {
  static ExcelValue result;
  if(variable_set[696] == 1) { return result;}
  result = _common55();
  variable_set[696] = 1;
  return result;
}

ExcelValue model_h53() {
  static ExcelValue result;
  if(variable_set[697] == 1) { return result;}
  result = _common57();
  variable_set[697] = 1;
  return result;
}

ExcelValue model_i53() {
  static ExcelValue result;
  if(variable_set[698] == 1) { return result;}
  result = _common59();
  variable_set[698] = 1;
  return result;
}

ExcelValue model_j53() {
  static ExcelValue result;
  if(variable_set[699] == 1) { return result;}
  result = _common61();
  variable_set[699] = 1;
  return result;
}

ExcelValue model_k53() {
  static ExcelValue result;
  if(variable_set[700] == 1) { return result;}
  result = _common63();
  variable_set[700] = 1;
  return result;
}

ExcelValue model_l53() {
  static ExcelValue result;
  if(variable_set[701] == 1) { return result;}
  result = _common65();
  variable_set[701] = 1;
  return result;
}

ExcelValue model_m53() {
  static ExcelValue result;
  if(variable_set[702] == 1) { return result;}
  result = divide(multiply(model_m51(),_common21()),C34);
  variable_set[702] = 1;
  return result;
}

ExcelValue model_n53() {
  static ExcelValue result;
  if(variable_set[703] == 1) { return result;}
  result = divide(multiply(model_n51(),_common22()),C34);
  variable_set[703] = 1;
  return result;
}

ExcelValue model_o53() {
  static ExcelValue result;
  if(variable_set[704] == 1) { return result;}
  result = divide(multiply(model_o51(),_common23()),C34);
  variable_set[704] = 1;
  return result;
}

ExcelValue model_p53() {
  static ExcelValue result;
  if(variable_set[705] == 1) { return result;}
  result = divide(multiply(model_p51(),_common24()),C34);
  variable_set[705] = 1;
  return result;
}

ExcelValue model_q53() {
  static ExcelValue result;
  if(variable_set[706] == 1) { return result;}
  result = divide(multiply(model_q51(),_common25()),C34);
  variable_set[706] = 1;
  return result;
}

ExcelValue model_r53() {
  static ExcelValue result;
  if(variable_set[707] == 1) { return result;}
  result = _common67();
  variable_set[707] = 1;
  return result;
}

ExcelValue model_s53() {
  static ExcelValue result;
  if(variable_set[708] == 1) { return result;}
  result = _common69();
  variable_set[708] = 1;
  return result;
}

ExcelValue model_t53() {
  static ExcelValue result;
  if(variable_set[709] == 1) { return result;}
  result = _common2();
  variable_set[709] = 1;
  return result;
}

ExcelValue model_u53() {
  static ExcelValue result;
  if(variable_set[710] == 1) { return result;}
  result = _common71();
  variable_set[710] = 1;
  return result;
}

ExcelValue model_v53() {
  static ExcelValue result;
  if(variable_set[711] == 1) { return result;}
  result = _common73();
  variable_set[711] = 1;
  return result;
}

ExcelValue model_w53() {
  static ExcelValue result;
  if(variable_set[712] == 1) { return result;}
  result = _common75();
  variable_set[712] = 1;
  return result;
}

ExcelValue model_x53() {
  static ExcelValue result;
  if(variable_set[713] == 1) { return result;}
  result = _common77();
  variable_set[713] = 1;
  return result;
}

ExcelValue model_y53() {
  static ExcelValue result;
  if(variable_set[714] == 1) { return result;}
  result = _common79();
  variable_set[714] = 1;
  return result;
}

ExcelValue model_z53() {
  static ExcelValue result;
  if(variable_set[715] == 1) { return result;}
  result = _common81();
  variable_set[715] = 1;
  return result;
}

ExcelValue model_aa53() {
  static ExcelValue result;
  if(variable_set[716] == 1) { return result;}
  result = _common83();
  variable_set[716] = 1;
  return result;
}

ExcelValue model_ab53() {
  static ExcelValue result;
  if(variable_set[717] == 1) { return result;}
  result = _common85();
  variable_set[717] = 1;
  return result;
}

ExcelValue model_ac53() {
  static ExcelValue result;
  if(variable_set[718] == 1) { return result;}
  result = _common87();
  variable_set[718] = 1;
  return result;
}

ExcelValue model_ad53() {
  static ExcelValue result;
  if(variable_set[719] == 1) { return result;}
  result = _common89();
  variable_set[719] = 1;
  return result;
}

ExcelValue model_ae53() {
  static ExcelValue result;
  if(variable_set[720] == 1) { return result;}
  result = _common91();
  variable_set[720] = 1;
  return result;
}

ExcelValue model_af53() {
  static ExcelValue result;
  if(variable_set[721] == 1) { return result;}
  result = _common93();
  variable_set[721] = 1;
  return result;
}

ExcelValue model_ag53() {
  static ExcelValue result;
  if(variable_set[722] == 1) { return result;}
  result = _common95();
  variable_set[722] = 1;
  return result;
}

ExcelValue model_ah53() {
  static ExcelValue result;
  if(variable_set[723] == 1) { return result;}
  result = _common97();
  variable_set[723] = 1;
  return result;
}

ExcelValue model_ai53() {
  static ExcelValue result;
  if(variable_set[724] == 1) { return result;}
  result = _common99();
  variable_set[724] = 1;
  return result;
}

ExcelValue model_aj53() {
  static ExcelValue result;
  if(variable_set[725] == 1) { return result;}
  result = _common101();
  variable_set[725] = 1;
  return result;
}

ExcelValue model_ak53() {
  static ExcelValue result;
  if(variable_set[726] == 1) { return result;}
  result = _common103();
  variable_set[726] = 1;
  return result;
}

ExcelValue model_al53() {
  static ExcelValue result;
  if(variable_set[727] == 1) { return result;}
  result = _common105();
  variable_set[727] = 1;
  return result;
}

ExcelValue model_am53() {
  static ExcelValue result;
  if(variable_set[728] == 1) { return result;}
  result = _common107();
  variable_set[728] = 1;
  return result;
}

ExcelValue model_an53() {
  static ExcelValue result;
  if(variable_set[729] == 1) { return result;}
  result = _common7();
  variable_set[729] = 1;
  return result;
}

ExcelValue model_c54() {
  static ExcelValue result;
  if(variable_set[730] == 1) { return result;}
  result = C35;
  variable_set[730] = 1;
  return result;
}

ExcelValue model_d54() {
  static ExcelValue result;
  if(variable_set[731] == 1) { return result;}
  result = _common109();
  variable_set[731] = 1;
  return result;
}

ExcelValue model_e54() {
  static ExcelValue result;
  if(variable_set[732] == 1) { return result;}
  result = _common110();
  variable_set[732] = 1;
  return result;
}

ExcelValue model_f54() {
  static ExcelValue result;
  if(variable_set[733] == 1) { return result;}
  result = _common111();
  variable_set[733] = 1;
  return result;
}

ExcelValue model_g54() {
  static ExcelValue result;
  if(variable_set[734] == 1) { return result;}
  result = _common112();
  variable_set[734] = 1;
  return result;
}

ExcelValue model_h54() {
  static ExcelValue result;
  if(variable_set[735] == 1) { return result;}
  result = _common113();
  variable_set[735] = 1;
  return result;
}

ExcelValue model_i54() {
  static ExcelValue result;
  if(variable_set[736] == 1) { return result;}
  result = _common114();
  variable_set[736] = 1;
  return result;
}

ExcelValue model_j54() {
  static ExcelValue result;
  if(variable_set[737] == 1) { return result;}
  result = _common115();
  variable_set[737] = 1;
  return result;
}

ExcelValue model_k54() {
  static ExcelValue result;
  if(variable_set[738] == 1) { return result;}
  result = multiply(model_j49(),C36);
  variable_set[738] = 1;
  return result;
}

ExcelValue model_l54() {
  static ExcelValue result;
  if(variable_set[739] == 1) { return result;}
  result = multiply(model_k74(),C36);
  variable_set[739] = 1;
  return result;
}

ExcelValue model_m54() {
  static ExcelValue result;
  if(variable_set[740] == 1) { return result;}
  result = multiply(model_l74(),C36);
  variable_set[740] = 1;
  return result;
}

ExcelValue model_n54() {
  static ExcelValue result;
  if(variable_set[741] == 1) { return result;}
  result = multiply(model_m74(),C36);
  variable_set[741] = 1;
  return result;
}

ExcelValue model_o54() {
  static ExcelValue result;
  if(variable_set[742] == 1) { return result;}
  result = multiply(model_n74(),C36);
  variable_set[742] = 1;
  return result;
}

ExcelValue model_p54() {
  static ExcelValue result;
  if(variable_set[743] == 1) { return result;}
  result = multiply(model_o74(),C36);
  variable_set[743] = 1;
  return result;
}

ExcelValue model_q54() {
  static ExcelValue result;
  if(variable_set[744] == 1) { return result;}
  result = multiply(model_p74(),C36);
  variable_set[744] = 1;
  return result;
}

ExcelValue model_r54() {
  static ExcelValue result;
  if(variable_set[745] == 1) { return result;}
  result = multiply(model_q74(),C36);
  variable_set[745] = 1;
  return result;
}

ExcelValue model_s54() {
  static ExcelValue result;
  if(variable_set[746] == 1) { return result;}
  result = multiply(model_r74(),C36);
  variable_set[746] = 1;
  return result;
}

ExcelValue model_t54() {
  static ExcelValue result;
  if(variable_set[747] == 1) { return result;}
  result = multiply(model_s74(),C36);
  variable_set[747] = 1;
  return result;
}

ExcelValue model_u54() {
  static ExcelValue result;
  if(variable_set[748] == 1) { return result;}
  result = multiply(model_t74(),C36);
  variable_set[748] = 1;
  return result;
}

ExcelValue model_v54() {
  static ExcelValue result;
  if(variable_set[749] == 1) { return result;}
  result = multiply(model_u74(),C36);
  variable_set[749] = 1;
  return result;
}

ExcelValue model_w54() {
  static ExcelValue result;
  if(variable_set[750] == 1) { return result;}
  result = multiply(model_v74(),C36);
  variable_set[750] = 1;
  return result;
}

ExcelValue model_x54() {
  static ExcelValue result;
  if(variable_set[751] == 1) { return result;}
  result = multiply(model_w74(),C36);
  variable_set[751] = 1;
  return result;
}

ExcelValue model_y54() {
  static ExcelValue result;
  if(variable_set[752] == 1) { return result;}
  result = multiply(model_x74(),C36);
  variable_set[752] = 1;
  return result;
}

ExcelValue model_z54() {
  static ExcelValue result;
  if(variable_set[753] == 1) { return result;}
  result = multiply(model_y74(),C36);
  variable_set[753] = 1;
  return result;
}

ExcelValue model_aa54() {
  static ExcelValue result;
  if(variable_set[754] == 1) { return result;}
  result = multiply(model_z74(),C36);
  variable_set[754] = 1;
  return result;
}

ExcelValue model_ab54() {
  static ExcelValue result;
  if(variable_set[755] == 1) { return result;}
  result = multiply(model_aa74(),C36);
  variable_set[755] = 1;
  return result;
}

ExcelValue model_ac54() {
  static ExcelValue result;
  if(variable_set[756] == 1) { return result;}
  result = multiply(model_ab74(),C36);
  variable_set[756] = 1;
  return result;
}

ExcelValue model_ad54() {
  static ExcelValue result;
  if(variable_set[757] == 1) { return result;}
  result = multiply(model_ac74(),C36);
  variable_set[757] = 1;
  return result;
}

ExcelValue model_ae54() {
  static ExcelValue result;
  if(variable_set[758] == 1) { return result;}
  result = multiply(model_ad74(),C36);
  variable_set[758] = 1;
  return result;
}

ExcelValue model_af54() {
  static ExcelValue result;
  if(variable_set[759] == 1) { return result;}
  result = multiply(model_ae74(),C36);
  variable_set[759] = 1;
  return result;
}

ExcelValue model_ag54() {
  static ExcelValue result;
  if(variable_set[760] == 1) { return result;}
  result = multiply(model_af74(),C36);
  variable_set[760] = 1;
  return result;
}

ExcelValue model_ah54() {
  static ExcelValue result;
  if(variable_set[761] == 1) { return result;}
  result = multiply(model_ag74(),C36);
  variable_set[761] = 1;
  return result;
}

ExcelValue model_ai54() {
  static ExcelValue result;
  if(variable_set[762] == 1) { return result;}
  result = multiply(model_ah74(),C36);
  variable_set[762] = 1;
  return result;
}

ExcelValue model_aj54() {
  static ExcelValue result;
  if(variable_set[763] == 1) { return result;}
  result = multiply(model_ai74(),C36);
  variable_set[763] = 1;
  return result;
}

ExcelValue model_ak54() {
  static ExcelValue result;
  if(variable_set[764] == 1) { return result;}
  result = multiply(model_aj74(),C36);
  variable_set[764] = 1;
  return result;
}

ExcelValue model_al54() {
  static ExcelValue result;
  if(variable_set[765] == 1) { return result;}
  result = multiply(model_ak74(),C36);
  variable_set[765] = 1;
  return result;
}

ExcelValue model_am54() {
  static ExcelValue result;
  if(variable_set[766] == 1) { return result;}
  result = multiply(model_al74(),C36);
  variable_set[766] = 1;
  return result;
}

ExcelValue model_an54() {
  static ExcelValue result;
  if(variable_set[767] == 1) { return result;}
  result = multiply(model_am74(),C36);
  variable_set[767] = 1;
  return result;
}

ExcelValue model_b55() {
  static ExcelValue result;
  if(variable_set[768] == 1) { return result;}
  result = _common116();
  variable_set[768] = 1;
  return result;
}

ExcelValue model_c55() {
  static ExcelValue result;
  if(variable_set[769] == 1) { return result;}
  result = _common116();
  variable_set[769] = 1;
  return result;
}

ExcelValue model_d55() {
  static ExcelValue result;
  if(variable_set[770] == 1) { return result;}
  result = add(_common118(),_common109());
  variable_set[770] = 1;
  return result;
}

ExcelValue model_e55() {
  static ExcelValue result;
  if(variable_set[771] == 1) { return result;}
  result = add(_common119(),_common110());
  variable_set[771] = 1;
  return result;
}

ExcelValue model_f55() {
  static ExcelValue result;
  if(variable_set[772] == 1) { return result;}
  result = add(_common120(),_common111());
  variable_set[772] = 1;
  return result;
}

ExcelValue model_g55() {
  static ExcelValue result;
  if(variable_set[773] == 1) { return result;}
  result = add(_common121(),_common112());
  variable_set[773] = 1;
  return result;
}

ExcelValue model_h55() {
  static ExcelValue result;
  if(variable_set[774] == 1) { return result;}
  result = add(_common122(),_common113());
  variable_set[774] = 1;
  return result;
}

ExcelValue model_i55() {
  static ExcelValue result;
  if(variable_set[775] == 1) { return result;}
  result = add(_common123(),_common114());
  variable_set[775] = 1;
  return result;
}

ExcelValue model_j55() {
  static ExcelValue result;
  if(variable_set[776] == 1) { return result;}
  result = add(_common124(),_common115());
  variable_set[776] = 1;
  return result;
}

ExcelValue model_k55() {
  static ExcelValue result;
  if(variable_set[777] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_j55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_k64(),subtract(model_j49(),model_k54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_j55(),C37),C37};
  result = excel_if(more_than(model_k47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[777] = 1;
  return result;
}

ExcelValue model_l55() {
  static ExcelValue result;
  if(variable_set[778] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_k55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_l64(),subtract(model_k74(),model_l54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_k55(),C37),C37};
  result = excel_if(more_than(model_l47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[778] = 1;
  return result;
}

ExcelValue model_m55() {
  static ExcelValue result;
  if(variable_set[779] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_l55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_m64(),subtract(model_l74(),model_m54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_l55(),C37),C37};
  result = excel_if(more_than(model_m47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[779] = 1;
  return result;
}

ExcelValue model_n55() {
  static ExcelValue result;
  if(variable_set[780] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_m55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_n64(),subtract(model_m74(),model_n54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_m55(),C37),C37};
  result = excel_if(more_than(model_n47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[780] = 1;
  return result;
}

ExcelValue model_o55() {
  static ExcelValue result;
  if(variable_set[781] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_n55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_o64(),subtract(model_n74(),model_o54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_n55(),C37),C37};
  result = excel_if(more_than(model_o47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[781] = 1;
  return result;
}

ExcelValue model_p55() {
  static ExcelValue result;
  if(variable_set[782] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_o55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_p64(),subtract(model_o74(),model_p54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_o55(),C37),C37};
  result = excel_if(more_than(model_p47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[782] = 1;
  return result;
}

ExcelValue model_q55() {
  static ExcelValue result;
  if(variable_set[783] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_p55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_q64(),subtract(model_p74(),model_q54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_p55(),C37),C37};
  result = excel_if(more_than(model_q47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[783] = 1;
  return result;
}

ExcelValue model_r55() {
  static ExcelValue result;
  if(variable_set[784] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_q55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_r64(),subtract(model_q74(),model_r54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_q55(),C37),C37};
  result = excel_if(more_than(model_r47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[784] = 1;
  return result;
}

ExcelValue model_s55() {
  static ExcelValue result;
  if(variable_set[785] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_r55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_s64(),subtract(model_r74(),model_s54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_r55(),C37),C37};
  result = excel_if(more_than(model_s47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[785] = 1;
  return result;
}

ExcelValue model_t55() {
  static ExcelValue result;
  if(variable_set[786] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_s55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_t64(),subtract(model_s74(),model_t54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_s55(),C37),C37};
  result = excel_if(more_than(model_t47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[786] = 1;
  return result;
}

ExcelValue model_u55() {
  static ExcelValue result;
  if(variable_set[787] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_t55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_u64(),subtract(model_t74(),model_u54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_t55(),C37),C37};
  result = excel_if(more_than(model_u47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[787] = 1;
  return result;
}

ExcelValue model_v55() {
  static ExcelValue result;
  if(variable_set[788] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_u55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_v64(),subtract(model_u74(),model_v54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_u55(),C37),C37};
  result = excel_if(more_than(model_v47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[788] = 1;
  return result;
}

ExcelValue model_w55() {
  static ExcelValue result;
  if(variable_set[789] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_v55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_w64(),subtract(model_v74(),model_w54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_v55(),C37),C37};
  result = excel_if(more_than(model_w47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[789] = 1;
  return result;
}

ExcelValue model_x55() {
  static ExcelValue result;
  if(variable_set[790] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_w55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_x64(),subtract(model_w74(),model_x54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_w55(),C37),C37};
  result = excel_if(more_than(model_x47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[790] = 1;
  return result;
}

ExcelValue model_y55() {
  static ExcelValue result;
  if(variable_set[791] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_x55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_y64(),subtract(model_x74(),model_y54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_x55(),C37),C37};
  result = excel_if(more_than(model_y47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[791] = 1;
  return result;
}

ExcelValue model_z55() {
  static ExcelValue result;
  if(variable_set[792] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_y55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_z64(),subtract(model_y74(),model_z54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_y55(),C37),C37};
  result = excel_if(more_than(model_z47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[792] = 1;
  return result;
}

ExcelValue model_aa55() {
  static ExcelValue result;
  if(variable_set[793] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_z55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_aa64(),subtract(model_z74(),model_aa54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_z55(),C37),C37};
  result = excel_if(more_than(model_aa47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[793] = 1;
  return result;
}

ExcelValue model_ab55() {
  static ExcelValue result;
  if(variable_set[794] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_aa55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_ab64(),subtract(model_aa74(),model_ab54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_aa55(),C37),C37};
  result = excel_if(more_than(model_ab47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[794] = 1;
  return result;
}

ExcelValue model_ac55() {
  static ExcelValue result;
  if(variable_set[795] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_ab55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_ac64(),subtract(model_ab74(),model_ac54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_ab55(),C37),C37};
  result = excel_if(more_than(model_ac47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[795] = 1;
  return result;
}

ExcelValue model_ad55() {
  static ExcelValue result;
  if(variable_set[796] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_ac55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_ad64(),subtract(model_ac74(),model_ad54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_ac55(),C37),C37};
  result = excel_if(more_than(model_ad47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[796] = 1;
  return result;
}

ExcelValue model_ae55() {
  static ExcelValue result;
  if(variable_set[797] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_ad55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_ae64(),subtract(model_ad74(),model_ae54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_ad55(),C37),C37};
  result = excel_if(more_than(model_ae47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[797] = 1;
  return result;
}

ExcelValue model_af55() {
  static ExcelValue result;
  if(variable_set[798] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_ae55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_af64(),subtract(model_ae74(),model_af54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_ae55(),C37),C37};
  result = excel_if(more_than(model_af47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[798] = 1;
  return result;
}

ExcelValue model_ag55() {
  static ExcelValue result;
  if(variable_set[799] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_af55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_ag64(),subtract(model_af74(),model_ag54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_af55(),C37),C37};
  result = excel_if(more_than(model_ag47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[799] = 1;
  return result;
}

ExcelValue model_ah55() {
  static ExcelValue result;
  if(variable_set[800] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_ag55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_ah64(),subtract(model_ag74(),model_ah54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_ag55(),C37),C37};
  result = excel_if(more_than(model_ah47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[800] = 1;
  return result;
}

ExcelValue model_ai55() {
  static ExcelValue result;
  if(variable_set[801] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_ah55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_ai64(),subtract(model_ah74(),model_ai54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_ah55(),C37),C37};
  result = excel_if(more_than(model_ai47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[801] = 1;
  return result;
}

ExcelValue model_aj55() {
  static ExcelValue result;
  if(variable_set[802] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_ai55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_aj64(),subtract(model_ai74(),model_aj54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_ai55(),C37),C37};
  result = excel_if(more_than(model_aj47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[802] = 1;
  return result;
}

ExcelValue model_ak55() {
  static ExcelValue result;
  if(variable_set[803] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_aj55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_ak64(),subtract(model_aj74(),model_ak54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_aj55(),C37),C37};
  result = excel_if(more_than(model_ak47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[803] = 1;
  return result;
}

ExcelValue model_al55() {
  static ExcelValue result;
  if(variable_set[804] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_ak55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_al64(),subtract(model_ak74(),model_al54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_ak55(),C37),C37};
  result = excel_if(more_than(model_al47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[804] = 1;
  return result;
}

ExcelValue model_am55() {
  static ExcelValue result;
  if(variable_set[805] == 1) { return result;}
  ExcelValue array1[] = {multiply(model_al55(),C22),C6};
  ExcelValue array2[] = {model_b9(),subtract(model_am64(),subtract(model_al74(),model_am54()))};
  ExcelValue array0[] = {max(2, array1),min(2, array2)};
  ExcelValue array3[] = {multiply(model_al55(),C37),C37};
  result = excel_if(more_than(model_am47(),model_b8()),min(2, array0),max(2, array3));
  variable_set[805] = 1;
  return result;
}

ExcelValue model_an55() {
  static ExcelValue result;
  if(variable_set[806] == 1) { return result;}
  result = _common125();
  variable_set[806] = 1;
  return result;
}

ExcelValue model_c56() {
  static ExcelValue result;
  if(variable_set[807] == 1) { return result;}
  result = _common117();
  variable_set[807] = 1;
  return result;
}

ExcelValue model_d56() {
  static ExcelValue result;
  if(variable_set[808] == 1) { return result;}
  result = _common118();
  variable_set[808] = 1;
  return result;
}

ExcelValue model_e56() {
  static ExcelValue result;
  if(variable_set[809] == 1) { return result;}
  result = _common119();
  variable_set[809] = 1;
  return result;
}

ExcelValue model_f56() {
  static ExcelValue result;
  if(variable_set[810] == 1) { return result;}
  result = _common120();
  variable_set[810] = 1;
  return result;
}

ExcelValue model_g56() {
  static ExcelValue result;
  if(variable_set[811] == 1) { return result;}
  result = _common121();
  variable_set[811] = 1;
  return result;
}

ExcelValue model_h56() {
  static ExcelValue result;
  if(variable_set[812] == 1) { return result;}
  result = _common122();
  variable_set[812] = 1;
  return result;
}

ExcelValue model_i56() {
  static ExcelValue result;
  if(variable_set[813] == 1) { return result;}
  result = _common123();
  variable_set[813] = 1;
  return result;
}

ExcelValue model_j56() {
  static ExcelValue result;
  if(variable_set[814] == 1) { return result;}
  result = _common124();
  variable_set[814] = 1;
  return result;
}

ExcelValue model_k56() {
  static ExcelValue result;
  if(variable_set[815] == 1) { return result;}
  result = _common135();
  variable_set[815] = 1;
  return result;
}

ExcelValue model_l56() {
  static ExcelValue result;
  if(variable_set[816] == 1) { return result;}
  result = _common136();
  variable_set[816] = 1;
  return result;
}

ExcelValue model_m56() {
  static ExcelValue result;
  if(variable_set[817] == 1) { return result;}
  result = _common137();
  variable_set[817] = 1;
  return result;
}

ExcelValue model_n56() {
  static ExcelValue result;
  if(variable_set[818] == 1) { return result;}
  result = _common138();
  variable_set[818] = 1;
  return result;
}

ExcelValue model_o56() {
  static ExcelValue result;
  if(variable_set[819] == 1) { return result;}
  result = _common139();
  variable_set[819] = 1;
  return result;
}

ExcelValue model_p56() {
  static ExcelValue result;
  if(variable_set[820] == 1) { return result;}
  result = _common140();
  variable_set[820] = 1;
  return result;
}

ExcelValue model_q56() {
  static ExcelValue result;
  if(variable_set[821] == 1) { return result;}
  result = _common141();
  variable_set[821] = 1;
  return result;
}

ExcelValue model_r56() {
  static ExcelValue result;
  if(variable_set[822] == 1) { return result;}
  result = _common142();
  variable_set[822] = 1;
  return result;
}

ExcelValue model_s56() {
  static ExcelValue result;
  if(variable_set[823] == 1) { return result;}
  result = _common143();
  variable_set[823] = 1;
  return result;
}

ExcelValue model_t56() {
  static ExcelValue result;
  if(variable_set[824] == 1) { return result;}
  result = _common144();
  variable_set[824] = 1;
  return result;
}

ExcelValue model_u56() {
  static ExcelValue result;
  if(variable_set[825] == 1) { return result;}
  result = _common145();
  variable_set[825] = 1;
  return result;
}

ExcelValue model_v56() {
  static ExcelValue result;
  if(variable_set[826] == 1) { return result;}
  result = _common146();
  variable_set[826] = 1;
  return result;
}

ExcelValue model_w56() {
  static ExcelValue result;
  if(variable_set[827] == 1) { return result;}
  result = _common147();
  variable_set[827] = 1;
  return result;
}

ExcelValue model_x56() {
  static ExcelValue result;
  if(variable_set[828] == 1) { return result;}
  result = _common148();
  variable_set[828] = 1;
  return result;
}

ExcelValue model_y56() {
  static ExcelValue result;
  if(variable_set[829] == 1) { return result;}
  result = _common149();
  variable_set[829] = 1;
  return result;
}

ExcelValue model_z56() {
  static ExcelValue result;
  if(variable_set[830] == 1) { return result;}
  result = _common150();
  variable_set[830] = 1;
  return result;
}

ExcelValue model_aa56() {
  static ExcelValue result;
  if(variable_set[831] == 1) { return result;}
  result = _common151();
  variable_set[831] = 1;
  return result;
}

ExcelValue model_ab56() {
  static ExcelValue result;
  if(variable_set[832] == 1) { return result;}
  result = _common152();
  variable_set[832] = 1;
  return result;
}

ExcelValue model_ac56() {
  static ExcelValue result;
  if(variable_set[833] == 1) { return result;}
  result = _common153();
  variable_set[833] = 1;
  return result;
}

ExcelValue model_ad56() {
  static ExcelValue result;
  if(variable_set[834] == 1) { return result;}
  result = _common154();
  variable_set[834] = 1;
  return result;
}

ExcelValue model_ae56() {
  static ExcelValue result;
  if(variable_set[835] == 1) { return result;}
  result = _common155();
  variable_set[835] = 1;
  return result;
}

ExcelValue model_af56() {
  static ExcelValue result;
  if(variable_set[836] == 1) { return result;}
  result = _common156();
  variable_set[836] = 1;
  return result;
}

ExcelValue model_ag56() {
  static ExcelValue result;
  if(variable_set[837] == 1) { return result;}
  result = _common157();
  variable_set[837] = 1;
  return result;
}

ExcelValue model_ah56() {
  static ExcelValue result;
  if(variable_set[838] == 1) { return result;}
  result = _common158();
  variable_set[838] = 1;
  return result;
}

ExcelValue model_ai56() {
  static ExcelValue result;
  if(variable_set[839] == 1) { return result;}
  result = _common159();
  variable_set[839] = 1;
  return result;
}

ExcelValue model_aj56() {
  static ExcelValue result;
  if(variable_set[840] == 1) { return result;}
  result = _common160();
  variable_set[840] = 1;
  return result;
}

ExcelValue model_ak56() {
  static ExcelValue result;
  if(variable_set[841] == 1) { return result;}
  result = _common161();
  variable_set[841] = 1;
  return result;
}

ExcelValue model_al56() {
  static ExcelValue result;
  if(variable_set[842] == 1) { return result;}
  result = _common162();
  variable_set[842] = 1;
  return result;
}

ExcelValue model_am56() {
  static ExcelValue result;
  if(variable_set[843] == 1) { return result;}
  result = _common163();
  variable_set[843] = 1;
  return result;
}

ExcelValue model_an56() {
  static ExcelValue result;
  if(variable_set[844] == 1) { return result;}
  result = _common164();
  variable_set[844] = 1;
  return result;
}

static ExcelValue model_d59() {
  static ExcelValue result;
  if(variable_set[845] == 1) { return result;}
  result = excel_if(TRUE,C7,multiply(C7,C38));
  variable_set[845] = 1;
  return result;
}

static ExcelValue model_e59() {
  static ExcelValue result;
  if(variable_set[846] == 1) { return result;}
  result = excel_if(_common165(),C7,multiply(model_d59(),C38));
  variable_set[846] = 1;
  return result;
}

static ExcelValue model_f59() {
  static ExcelValue result;
  if(variable_set[847] == 1) { return result;}
  result = excel_if(_common166(),C7,multiply(model_e59(),C38));
  variable_set[847] = 1;
  return result;
}

static ExcelValue model_g59() {
  static ExcelValue result;
  if(variable_set[848] == 1) { return result;}
  result = excel_if(_common167(),C7,multiply(model_f59(),C38));
  variable_set[848] = 1;
  return result;
}

static ExcelValue model_h59() {
  static ExcelValue result;
  if(variable_set[849] == 1) { return result;}
  result = excel_if(_common168(),C7,multiply(model_g59(),C38));
  variable_set[849] = 1;
  return result;
}

static ExcelValue model_i59() {
  static ExcelValue result;
  if(variable_set[850] == 1) { return result;}
  result = excel_if(_common169(),C7,multiply(model_h59(),C38));
  variable_set[850] = 1;
  return result;
}

static ExcelValue model_j59() {
  static ExcelValue result;
  if(variable_set[851] == 1) { return result;}
  result = excel_if(_common170(),C7,multiply(model_i59(),C38));
  variable_set[851] = 1;
  return result;
}

static ExcelValue model_k59() {
  static ExcelValue result;
  if(variable_set[852] == 1) { return result;}
  result = excel_if(_common171(),C7,multiply(model_j59(),C38));
  variable_set[852] = 1;
  return result;
}

static ExcelValue model_l59() {
  static ExcelValue result;
  if(variable_set[853] == 1) { return result;}
  result = excel_if(_common172(),C7,multiply(model_k59(),C38));
  variable_set[853] = 1;
  return result;
}

static ExcelValue model_m59() {
  static ExcelValue result;
  if(variable_set[854] == 1) { return result;}
  result = excel_if(_common173(),C7,multiply(model_l59(),C38));
  variable_set[854] = 1;
  return result;
}

static ExcelValue model_n59() {
  static ExcelValue result;
  if(variable_set[855] == 1) { return result;}
  result = excel_if(_common174(),C7,multiply(model_m59(),C38));
  variable_set[855] = 1;
  return result;
}

static ExcelValue model_o59() {
  static ExcelValue result;
  if(variable_set[856] == 1) { return result;}
  result = excel_if(_common175(),C7,multiply(model_n59(),C38));
  variable_set[856] = 1;
  return result;
}

static ExcelValue model_p59() {
  static ExcelValue result;
  if(variable_set[857] == 1) { return result;}
  result = excel_if(_common176(),C7,multiply(model_o59(),C38));
  variable_set[857] = 1;
  return result;
}

static ExcelValue model_q59() {
  static ExcelValue result;
  if(variable_set[858] == 1) { return result;}
  result = excel_if(_common177(),C7,multiply(model_p59(),C38));
  variable_set[858] = 1;
  return result;
}

static ExcelValue model_r59() {
  static ExcelValue result;
  if(variable_set[859] == 1) { return result;}
  result = excel_if(_common178(),C7,multiply(model_q59(),C38));
  variable_set[859] = 1;
  return result;
}

static ExcelValue model_s59() {
  static ExcelValue result;
  if(variable_set[860] == 1) { return result;}
  result = excel_if(_common179(),C7,multiply(model_r59(),C38));
  variable_set[860] = 1;
  return result;
}

static ExcelValue model_t59() {
  static ExcelValue result;
  if(variable_set[861] == 1) { return result;}
  result = excel_if(_common180(),C7,multiply(model_s59(),C38));
  variable_set[861] = 1;
  return result;
}

static ExcelValue model_u59() {
  static ExcelValue result;
  if(variable_set[862] == 1) { return result;}
  result = excel_if(_common181(),C7,multiply(model_t59(),C38));
  variable_set[862] = 1;
  return result;
}

static ExcelValue model_v59() {
  static ExcelValue result;
  if(variable_set[863] == 1) { return result;}
  result = excel_if(_common182(),C7,multiply(model_u59(),C38));
  variable_set[863] = 1;
  return result;
}

static ExcelValue model_w59() {
  static ExcelValue result;
  if(variable_set[864] == 1) { return result;}
  result = excel_if(_common183(),C7,multiply(model_v59(),C38));
  variable_set[864] = 1;
  return result;
}

static ExcelValue model_x59() {
  static ExcelValue result;
  if(variable_set[865] == 1) { return result;}
  result = excel_if(_common184(),C7,multiply(model_w59(),C38));
  variable_set[865] = 1;
  return result;
}

static ExcelValue model_y59() {
  static ExcelValue result;
  if(variable_set[866] == 1) { return result;}
  result = excel_if(_common185(),C7,multiply(model_x59(),C38));
  variable_set[866] = 1;
  return result;
}

static ExcelValue model_z59() {
  static ExcelValue result;
  if(variable_set[867] == 1) { return result;}
  result = excel_if(_common186(),C7,multiply(model_y59(),C38));
  variable_set[867] = 1;
  return result;
}

static ExcelValue model_aa59() {
  static ExcelValue result;
  if(variable_set[868] == 1) { return result;}
  result = excel_if(_common187(),C7,multiply(model_z59(),C38));
  variable_set[868] = 1;
  return result;
}

static ExcelValue model_ab59() {
  static ExcelValue result;
  if(variable_set[869] == 1) { return result;}
  result = excel_if(_common188(),C7,multiply(model_aa59(),C38));
  variable_set[869] = 1;
  return result;
}

static ExcelValue model_ac59() {
  static ExcelValue result;
  if(variable_set[870] == 1) { return result;}
  result = excel_if(_common189(),C7,multiply(model_ab59(),C38));
  variable_set[870] = 1;
  return result;
}

static ExcelValue model_ad59() {
  static ExcelValue result;
  if(variable_set[871] == 1) { return result;}
  result = excel_if(_common190(),C7,multiply(model_ac59(),C38));
  variable_set[871] = 1;
  return result;
}

static ExcelValue model_ae59() {
  static ExcelValue result;
  if(variable_set[872] == 1) { return result;}
  result = excel_if(_common191(),C7,multiply(model_ad59(),C38));
  variable_set[872] = 1;
  return result;
}

static ExcelValue model_af59() {
  static ExcelValue result;
  if(variable_set[873] == 1) { return result;}
  result = excel_if(_common192(),C7,multiply(model_ae59(),C38));
  variable_set[873] = 1;
  return result;
}

static ExcelValue model_ag59() {
  static ExcelValue result;
  if(variable_set[874] == 1) { return result;}
  result = excel_if(_common193(),C7,multiply(model_af59(),C38));
  variable_set[874] = 1;
  return result;
}

static ExcelValue model_ah59() {
  static ExcelValue result;
  if(variable_set[875] == 1) { return result;}
  result = excel_if(_common194(),C7,multiply(model_ag59(),C38));
  variable_set[875] = 1;
  return result;
}

static ExcelValue model_ai59() {
  static ExcelValue result;
  if(variable_set[876] == 1) { return result;}
  result = excel_if(_common195(),C7,multiply(model_ah59(),C38));
  variable_set[876] = 1;
  return result;
}

static ExcelValue model_aj59() {
  static ExcelValue result;
  if(variable_set[877] == 1) { return result;}
  result = excel_if(_common196(),C7,multiply(model_ai59(),C38));
  variable_set[877] = 1;
  return result;
}

static ExcelValue model_ak59() {
  static ExcelValue result;
  if(variable_set[878] == 1) { return result;}
  result = excel_if(_common197(),C7,multiply(model_aj59(),C38));
  variable_set[878] = 1;
  return result;
}

static ExcelValue model_al59() {
  static ExcelValue result;
  if(variable_set[879] == 1) { return result;}
  result = excel_if(_common198(),C7,multiply(model_ak59(),C38));
  variable_set[879] = 1;
  return result;
}

static ExcelValue model_am59() {
  static ExcelValue result;
  if(variable_set[880] == 1) { return result;}
  result = excel_if(_common199(),C7,multiply(model_al59(),C38));
  variable_set[880] = 1;
  return result;
}

static ExcelValue model_d60() {
  static ExcelValue result;
  if(variable_set[881] == 1) { return result;}
  result = excel_if(TRUE,C22,multiply(C22,C39));
  variable_set[881] = 1;
  return result;
}

static ExcelValue model_e60() {
  static ExcelValue result;
  if(variable_set[882] == 1) { return result;}
  result = excel_if(_common165(),C22,multiply(model_d60(),C39));
  variable_set[882] = 1;
  return result;
}

static ExcelValue model_f60() {
  static ExcelValue result;
  if(variable_set[883] == 1) { return result;}
  result = excel_if(_common166(),C22,multiply(model_e60(),C39));
  variable_set[883] = 1;
  return result;
}

static ExcelValue model_g60() {
  static ExcelValue result;
  if(variable_set[884] == 1) { return result;}
  result = excel_if(_common167(),C22,multiply(model_f60(),C39));
  variable_set[884] = 1;
  return result;
}

static ExcelValue model_h60() {
  static ExcelValue result;
  if(variable_set[885] == 1) { return result;}
  result = excel_if(_common168(),C22,multiply(model_g60(),C39));
  variable_set[885] = 1;
  return result;
}

static ExcelValue model_i60() {
  static ExcelValue result;
  if(variable_set[886] == 1) { return result;}
  result = excel_if(_common169(),C22,multiply(model_h60(),C39));
  variable_set[886] = 1;
  return result;
}

static ExcelValue model_j60() {
  static ExcelValue result;
  if(variable_set[887] == 1) { return result;}
  result = excel_if(_common170(),C22,multiply(model_i60(),C39));
  variable_set[887] = 1;
  return result;
}

static ExcelValue model_k60() {
  static ExcelValue result;
  if(variable_set[888] == 1) { return result;}
  result = excel_if(_common171(),C22,multiply(model_j60(),C39));
  variable_set[888] = 1;
  return result;
}

static ExcelValue model_l60() {
  static ExcelValue result;
  if(variable_set[889] == 1) { return result;}
  result = excel_if(_common172(),C22,multiply(model_k60(),C39));
  variable_set[889] = 1;
  return result;
}

static ExcelValue model_m60() {
  static ExcelValue result;
  if(variable_set[890] == 1) { return result;}
  result = excel_if(_common173(),C22,multiply(model_l60(),C39));
  variable_set[890] = 1;
  return result;
}

static ExcelValue model_n60() {
  static ExcelValue result;
  if(variable_set[891] == 1) { return result;}
  result = excel_if(_common174(),C22,multiply(model_m60(),C39));
  variable_set[891] = 1;
  return result;
}

static ExcelValue model_o60() {
  static ExcelValue result;
  if(variable_set[892] == 1) { return result;}
  result = excel_if(_common175(),C22,multiply(model_n60(),C39));
  variable_set[892] = 1;
  return result;
}

static ExcelValue model_p60() {
  static ExcelValue result;
  if(variable_set[893] == 1) { return result;}
  result = excel_if(_common176(),C22,multiply(model_o60(),C39));
  variable_set[893] = 1;
  return result;
}

static ExcelValue model_q60() {
  static ExcelValue result;
  if(variable_set[894] == 1) { return result;}
  result = excel_if(_common177(),C22,multiply(model_p60(),C39));
  variable_set[894] = 1;
  return result;
}

static ExcelValue model_r60() {
  static ExcelValue result;
  if(variable_set[895] == 1) { return result;}
  result = excel_if(_common178(),C22,multiply(model_q60(),C39));
  variable_set[895] = 1;
  return result;
}

static ExcelValue model_s60() {
  static ExcelValue result;
  if(variable_set[896] == 1) { return result;}
  result = excel_if(_common179(),C22,multiply(model_r60(),C39));
  variable_set[896] = 1;
  return result;
}

static ExcelValue model_t60() {
  static ExcelValue result;
  if(variable_set[897] == 1) { return result;}
  result = excel_if(_common180(),C22,multiply(model_s60(),C39));
  variable_set[897] = 1;
  return result;
}

static ExcelValue model_u60() {
  static ExcelValue result;
  if(variable_set[898] == 1) { return result;}
  result = excel_if(_common181(),C22,multiply(model_t60(),C39));
  variable_set[898] = 1;
  return result;
}

static ExcelValue model_v60() {
  static ExcelValue result;
  if(variable_set[899] == 1) { return result;}
  result = excel_if(_common182(),C22,multiply(model_u60(),C39));
  variable_set[899] = 1;
  return result;
}

static ExcelValue model_w60() {
  static ExcelValue result;
  if(variable_set[900] == 1) { return result;}
  result = excel_if(_common183(),C22,multiply(model_v60(),C39));
  variable_set[900] = 1;
  return result;
}

static ExcelValue model_x60() {
  static ExcelValue result;
  if(variable_set[901] == 1) { return result;}
  result = excel_if(_common184(),C22,multiply(model_w60(),C39));
  variable_set[901] = 1;
  return result;
}

static ExcelValue model_y60() {
  static ExcelValue result;
  if(variable_set[902] == 1) { return result;}
  result = excel_if(_common185(),C22,multiply(model_x60(),C39));
  variable_set[902] = 1;
  return result;
}

static ExcelValue model_z60() {
  static ExcelValue result;
  if(variable_set[903] == 1) { return result;}
  result = excel_if(_common186(),C22,multiply(model_y60(),C39));
  variable_set[903] = 1;
  return result;
}

static ExcelValue model_aa60() {
  static ExcelValue result;
  if(variable_set[904] == 1) { return result;}
  result = excel_if(_common187(),C22,multiply(model_z60(),C39));
  variable_set[904] = 1;
  return result;
}

static ExcelValue model_ab60() {
  static ExcelValue result;
  if(variable_set[905] == 1) { return result;}
  result = excel_if(_common188(),C22,multiply(model_aa60(),C39));
  variable_set[905] = 1;
  return result;
}

static ExcelValue model_ac60() {
  static ExcelValue result;
  if(variable_set[906] == 1) { return result;}
  result = excel_if(_common189(),C22,multiply(model_ab60(),C39));
  variable_set[906] = 1;
  return result;
}

static ExcelValue model_ad60() {
  static ExcelValue result;
  if(variable_set[907] == 1) { return result;}
  result = excel_if(_common190(),C22,multiply(model_ac60(),C39));
  variable_set[907] = 1;
  return result;
}

static ExcelValue model_ae60() {
  static ExcelValue result;
  if(variable_set[908] == 1) { return result;}
  result = excel_if(_common191(),C22,multiply(model_ad60(),C39));
  variable_set[908] = 1;
  return result;
}

static ExcelValue model_af60() {
  static ExcelValue result;
  if(variable_set[909] == 1) { return result;}
  result = excel_if(_common192(),C22,multiply(model_ae60(),C39));
  variable_set[909] = 1;
  return result;
}

static ExcelValue model_ag60() {
  static ExcelValue result;
  if(variable_set[910] == 1) { return result;}
  result = excel_if(_common193(),C22,multiply(model_af60(),C39));
  variable_set[910] = 1;
  return result;
}

static ExcelValue model_ah60() {
  static ExcelValue result;
  if(variable_set[911] == 1) { return result;}
  result = excel_if(_common194(),C22,multiply(model_ag60(),C39));
  variable_set[911] = 1;
  return result;
}

static ExcelValue model_ai60() {
  static ExcelValue result;
  if(variable_set[912] == 1) { return result;}
  result = excel_if(_common195(),C22,multiply(model_ah60(),C39));
  variable_set[912] = 1;
  return result;
}

static ExcelValue model_aj60() {
  static ExcelValue result;
  if(variable_set[913] == 1) { return result;}
  result = excel_if(_common196(),C22,multiply(model_ai60(),C39));
  variable_set[913] = 1;
  return result;
}

static ExcelValue model_ak60() {
  static ExcelValue result;
  if(variable_set[914] == 1) { return result;}
  result = excel_if(_common197(),C22,multiply(model_aj60(),C39));
  variable_set[914] = 1;
  return result;
}

static ExcelValue model_al60() {
  static ExcelValue result;
  if(variable_set[915] == 1) { return result;}
  result = excel_if(_common198(),C22,multiply(model_ak60(),C39));
  variable_set[915] = 1;
  return result;
}

static ExcelValue model_am60() {
  static ExcelValue result;
  if(variable_set[916] == 1) { return result;}
  result = excel_if(_common199(),C22,multiply(model_al60(),C39));
  variable_set[916] = 1;
  return result;
}

static ExcelValue model_c63() {
  static ExcelValue result;
  if(variable_set[917] == 1) { return result;}
  result = multiply(C25,C7);
  variable_set[917] = 1;
  return result;
}

static ExcelValue model_d63() {
  static ExcelValue result;
  if(variable_set[918] == 1) { return result;}
  result = multiply(model_d48(),model_d59());
  variable_set[918] = 1;
  return result;
}

static ExcelValue model_e63() {
  static ExcelValue result;
  if(variable_set[919] == 1) { return result;}
  result = multiply(model_e48(),model_e59());
  variable_set[919] = 1;
  return result;
}

static ExcelValue model_f63() {
  static ExcelValue result;
  if(variable_set[920] == 1) { return result;}
  result = multiply(model_f48(),model_f59());
  variable_set[920] = 1;
  return result;
}

static ExcelValue model_g63() {
  static ExcelValue result;
  if(variable_set[921] == 1) { return result;}
  result = multiply(model_g48(),model_g59());
  variable_set[921] = 1;
  return result;
}

static ExcelValue model_h63() {
  static ExcelValue result;
  if(variable_set[922] == 1) { return result;}
  result = multiply(model_h48(),model_h59());
  variable_set[922] = 1;
  return result;
}

static ExcelValue model_i63() {
  static ExcelValue result;
  if(variable_set[923] == 1) { return result;}
  result = multiply(model_i48(),model_i59());
  variable_set[923] = 1;
  return result;
}

static ExcelValue model_j63() {
  static ExcelValue result;
  if(variable_set[924] == 1) { return result;}
  result = multiply(model_j48(),model_j59());
  variable_set[924] = 1;
  return result;
}

static ExcelValue model_k63() {
  static ExcelValue result;
  if(variable_set[925] == 1) { return result;}
  result = multiply(model_k48(),model_k59());
  variable_set[925] = 1;
  return result;
}

static ExcelValue model_l63() {
  static ExcelValue result;
  if(variable_set[926] == 1) { return result;}
  result = multiply(model_l48(),model_l59());
  variable_set[926] = 1;
  return result;
}

static ExcelValue model_m63() {
  static ExcelValue result;
  if(variable_set[927] == 1) { return result;}
  result = multiply(model_m48(),model_m59());
  variable_set[927] = 1;
  return result;
}

static ExcelValue model_n63() {
  static ExcelValue result;
  if(variable_set[928] == 1) { return result;}
  result = multiply(model_n48(),model_n59());
  variable_set[928] = 1;
  return result;
}

static ExcelValue model_o63() {
  static ExcelValue result;
  if(variable_set[929] == 1) { return result;}
  result = multiply(model_o48(),model_o59());
  variable_set[929] = 1;
  return result;
}

static ExcelValue model_p63() {
  static ExcelValue result;
  if(variable_set[930] == 1) { return result;}
  result = multiply(model_p48(),model_p59());
  variable_set[930] = 1;
  return result;
}

static ExcelValue model_q63() {
  static ExcelValue result;
  if(variable_set[931] == 1) { return result;}
  result = multiply(model_q48(),model_q59());
  variable_set[931] = 1;
  return result;
}

static ExcelValue model_r63() {
  static ExcelValue result;
  if(variable_set[932] == 1) { return result;}
  result = multiply(model_r48(),model_r59());
  variable_set[932] = 1;
  return result;
}

static ExcelValue model_s63() {
  static ExcelValue result;
  if(variable_set[933] == 1) { return result;}
  result = multiply(model_s48(),model_s59());
  variable_set[933] = 1;
  return result;
}

static ExcelValue model_t63() {
  static ExcelValue result;
  if(variable_set[934] == 1) { return result;}
  result = multiply(model_t48(),model_t59());
  variable_set[934] = 1;
  return result;
}

static ExcelValue model_u63() {
  static ExcelValue result;
  if(variable_set[935] == 1) { return result;}
  result = multiply(model_u48(),model_u59());
  variable_set[935] = 1;
  return result;
}

static ExcelValue model_v63() {
  static ExcelValue result;
  if(variable_set[936] == 1) { return result;}
  result = multiply(model_v48(),model_v59());
  variable_set[936] = 1;
  return result;
}

static ExcelValue model_w63() {
  static ExcelValue result;
  if(variable_set[937] == 1) { return result;}
  result = multiply(model_w48(),model_w59());
  variable_set[937] = 1;
  return result;
}

static ExcelValue model_x63() {
  static ExcelValue result;
  if(variable_set[938] == 1) { return result;}
  result = multiply(model_x48(),model_x59());
  variable_set[938] = 1;
  return result;
}

static ExcelValue model_y63() {
  static ExcelValue result;
  if(variable_set[939] == 1) { return result;}
  result = multiply(model_y48(),model_y59());
  variable_set[939] = 1;
  return result;
}

static ExcelValue model_z63() {
  static ExcelValue result;
  if(variable_set[940] == 1) { return result;}
  result = multiply(model_z48(),model_z59());
  variable_set[940] = 1;
  return result;
}

static ExcelValue model_aa63() {
  static ExcelValue result;
  if(variable_set[941] == 1) { return result;}
  result = multiply(model_aa48(),model_aa59());
  variable_set[941] = 1;
  return result;
}

static ExcelValue model_ab63() {
  static ExcelValue result;
  if(variable_set[942] == 1) { return result;}
  result = multiply(model_ab48(),model_ab59());
  variable_set[942] = 1;
  return result;
}

static ExcelValue model_ac63() {
  static ExcelValue result;
  if(variable_set[943] == 1) { return result;}
  result = multiply(model_ac48(),model_ac59());
  variable_set[943] = 1;
  return result;
}

static ExcelValue model_ad63() {
  static ExcelValue result;
  if(variable_set[944] == 1) { return result;}
  result = multiply(model_ad48(),model_ad59());
  variable_set[944] = 1;
  return result;
}

static ExcelValue model_ae63() {
  static ExcelValue result;
  if(variable_set[945] == 1) { return result;}
  result = multiply(model_ae48(),model_ae59());
  variable_set[945] = 1;
  return result;
}

static ExcelValue model_af63() {
  static ExcelValue result;
  if(variable_set[946] == 1) { return result;}
  result = multiply(model_af48(),model_af59());
  variable_set[946] = 1;
  return result;
}

static ExcelValue model_ag63() {
  static ExcelValue result;
  if(variable_set[947] == 1) { return result;}
  result = multiply(model_ag48(),model_ag59());
  variable_set[947] = 1;
  return result;
}

static ExcelValue model_ah63() {
  static ExcelValue result;
  if(variable_set[948] == 1) { return result;}
  result = multiply(model_ah48(),model_ah59());
  variable_set[948] = 1;
  return result;
}

static ExcelValue model_ai63() {
  static ExcelValue result;
  if(variable_set[949] == 1) { return result;}
  result = multiply(model_ai48(),model_ai59());
  variable_set[949] = 1;
  return result;
}

static ExcelValue model_aj63() {
  static ExcelValue result;
  if(variable_set[950] == 1) { return result;}
  result = multiply(model_aj48(),model_aj59());
  variable_set[950] = 1;
  return result;
}

static ExcelValue model_ak63() {
  static ExcelValue result;
  if(variable_set[951] == 1) { return result;}
  result = multiply(model_ak48(),model_ak59());
  variable_set[951] = 1;
  return result;
}

static ExcelValue model_al63() {
  static ExcelValue result;
  if(variable_set[952] == 1) { return result;}
  result = multiply(model_al48(),model_al59());
  variable_set[952] = 1;
  return result;
}

static ExcelValue model_am63() {
  static ExcelValue result;
  if(variable_set[953] == 1) { return result;}
  result = multiply(model_am48(),model_am59());
  variable_set[953] = 1;
  return result;
}

static ExcelValue model_an63() {
  static ExcelValue result;
  if(variable_set[954] == 1) { return result;}
  result = multiply(model_an48(),excel_if(_common200(),C7,multiply(model_am59(),C38)));
  variable_set[954] = 1;
  return result;
}

static ExcelValue model_c64() {
  static ExcelValue result;
  if(variable_set[955] == 1) { return result;}
  result = multiply(C22,C25);
  variable_set[955] = 1;
  return result;
}

static ExcelValue model_d64() {
  static ExcelValue result;
  if(variable_set[956] == 1) { return result;}
  result = multiply(model_d60(),model_d48());
  variable_set[956] = 1;
  return result;
}

static ExcelValue model_e64() {
  static ExcelValue result;
  if(variable_set[957] == 1) { return result;}
  result = multiply(model_e60(),model_e48());
  variable_set[957] = 1;
  return result;
}

static ExcelValue model_f64() {
  static ExcelValue result;
  if(variable_set[958] == 1) { return result;}
  result = multiply(model_f60(),model_f48());
  variable_set[958] = 1;
  return result;
}

static ExcelValue model_g64() {
  static ExcelValue result;
  if(variable_set[959] == 1) { return result;}
  result = multiply(model_g60(),model_g48());
  variable_set[959] = 1;
  return result;
}

static ExcelValue model_h64() {
  static ExcelValue result;
  if(variable_set[960] == 1) { return result;}
  result = multiply(model_h60(),model_h48());
  variable_set[960] = 1;
  return result;
}

static ExcelValue model_i64() {
  static ExcelValue result;
  if(variable_set[961] == 1) { return result;}
  result = multiply(model_i60(),model_i48());
  variable_set[961] = 1;
  return result;
}

static ExcelValue model_j64() {
  static ExcelValue result;
  if(variable_set[962] == 1) { return result;}
  result = multiply(model_j60(),model_j48());
  variable_set[962] = 1;
  return result;
}

static ExcelValue model_k64() {
  static ExcelValue result;
  if(variable_set[963] == 1) { return result;}
  result = multiply(model_k60(),model_k48());
  variable_set[963] = 1;
  return result;
}

static ExcelValue model_l64() {
  static ExcelValue result;
  if(variable_set[964] == 1) { return result;}
  result = multiply(model_l60(),model_l48());
  variable_set[964] = 1;
  return result;
}

static ExcelValue model_m64() {
  static ExcelValue result;
  if(variable_set[965] == 1) { return result;}
  result = multiply(model_m60(),model_m48());
  variable_set[965] = 1;
  return result;
}

static ExcelValue model_n64() {
  static ExcelValue result;
  if(variable_set[966] == 1) { return result;}
  result = multiply(model_n60(),model_n48());
  variable_set[966] = 1;
  return result;
}

static ExcelValue model_o64() {
  static ExcelValue result;
  if(variable_set[967] == 1) { return result;}
  result = multiply(model_o60(),model_o48());
  variable_set[967] = 1;
  return result;
}

static ExcelValue model_p64() {
  static ExcelValue result;
  if(variable_set[968] == 1) { return result;}
  result = multiply(model_p60(),model_p48());
  variable_set[968] = 1;
  return result;
}

static ExcelValue model_q64() {
  static ExcelValue result;
  if(variable_set[969] == 1) { return result;}
  result = multiply(model_q60(),model_q48());
  variable_set[969] = 1;
  return result;
}

static ExcelValue model_r64() {
  static ExcelValue result;
  if(variable_set[970] == 1) { return result;}
  result = multiply(model_r60(),model_r48());
  variable_set[970] = 1;
  return result;
}

static ExcelValue model_s64() {
  static ExcelValue result;
  if(variable_set[971] == 1) { return result;}
  result = multiply(model_s60(),model_s48());
  variable_set[971] = 1;
  return result;
}

static ExcelValue model_t64() {
  static ExcelValue result;
  if(variable_set[972] == 1) { return result;}
  result = multiply(model_t60(),model_t48());
  variable_set[972] = 1;
  return result;
}

static ExcelValue model_u64() {
  static ExcelValue result;
  if(variable_set[973] == 1) { return result;}
  result = multiply(model_u60(),model_u48());
  variable_set[973] = 1;
  return result;
}

static ExcelValue model_v64() {
  static ExcelValue result;
  if(variable_set[974] == 1) { return result;}
  result = multiply(model_v60(),model_v48());
  variable_set[974] = 1;
  return result;
}

static ExcelValue model_w64() {
  static ExcelValue result;
  if(variable_set[975] == 1) { return result;}
  result = multiply(model_w60(),model_w48());
  variable_set[975] = 1;
  return result;
}

static ExcelValue model_x64() {
  static ExcelValue result;
  if(variable_set[976] == 1) { return result;}
  result = multiply(model_x60(),model_x48());
  variable_set[976] = 1;
  return result;
}

static ExcelValue model_y64() {
  static ExcelValue result;
  if(variable_set[977] == 1) { return result;}
  result = multiply(model_y60(),model_y48());
  variable_set[977] = 1;
  return result;
}

static ExcelValue model_z64() {
  static ExcelValue result;
  if(variable_set[978] == 1) { return result;}
  result = multiply(model_z60(),model_z48());
  variable_set[978] = 1;
  return result;
}

static ExcelValue model_aa64() {
  static ExcelValue result;
  if(variable_set[979] == 1) { return result;}
  result = multiply(model_aa60(),model_aa48());
  variable_set[979] = 1;
  return result;
}

static ExcelValue model_ab64() {
  static ExcelValue result;
  if(variable_set[980] == 1) { return result;}
  result = multiply(model_ab60(),model_ab48());
  variable_set[980] = 1;
  return result;
}

static ExcelValue model_ac64() {
  static ExcelValue result;
  if(variable_set[981] == 1) { return result;}
  result = multiply(model_ac60(),model_ac48());
  variable_set[981] = 1;
  return result;
}

static ExcelValue model_ad64() {
  static ExcelValue result;
  if(variable_set[982] == 1) { return result;}
  result = multiply(model_ad60(),model_ad48());
  variable_set[982] = 1;
  return result;
}

static ExcelValue model_ae64() {
  static ExcelValue result;
  if(variable_set[983] == 1) { return result;}
  result = multiply(model_ae60(),model_ae48());
  variable_set[983] = 1;
  return result;
}

static ExcelValue model_af64() {
  static ExcelValue result;
  if(variable_set[984] == 1) { return result;}
  result = multiply(model_af60(),model_af48());
  variable_set[984] = 1;
  return result;
}

static ExcelValue model_ag64() {
  static ExcelValue result;
  if(variable_set[985] == 1) { return result;}
  result = multiply(model_ag60(),model_ag48());
  variable_set[985] = 1;
  return result;
}

static ExcelValue model_ah64() {
  static ExcelValue result;
  if(variable_set[986] == 1) { return result;}
  result = multiply(model_ah60(),model_ah48());
  variable_set[986] = 1;
  return result;
}

static ExcelValue model_ai64() {
  static ExcelValue result;
  if(variable_set[987] == 1) { return result;}
  result = multiply(model_ai60(),model_ai48());
  variable_set[987] = 1;
  return result;
}

static ExcelValue model_aj64() {
  static ExcelValue result;
  if(variable_set[988] == 1) { return result;}
  result = multiply(model_aj60(),model_aj48());
  variable_set[988] = 1;
  return result;
}

static ExcelValue model_ak64() {
  static ExcelValue result;
  if(variable_set[989] == 1) { return result;}
  result = multiply(model_ak60(),model_ak48());
  variable_set[989] = 1;
  return result;
}

static ExcelValue model_al64() {
  static ExcelValue result;
  if(variable_set[990] == 1) { return result;}
  result = multiply(model_al60(),model_al48());
  variable_set[990] = 1;
  return result;
}

static ExcelValue model_am64() {
  static ExcelValue result;
  if(variable_set[991] == 1) { return result;}
  result = multiply(model_am60(),model_am48());
  variable_set[991] = 1;
  return result;
}

static ExcelValue model_an64() {
  static ExcelValue result;
  if(variable_set[992] == 1) { return result;}
  result = multiply(excel_if(_common200(),C22,multiply(model_am60(),C39)),model_an48());
  variable_set[992] = 1;
  return result;
}

static ExcelValue model_b67() {
  static ExcelValue result;
  if(variable_set[993] == 1) { return result;}
  result = _common201();
  variable_set[993] = 1;
  return result;
}

static ExcelValue model_c67() {
  static ExcelValue result;
  if(variable_set[994] == 1) { return result;}
  result = _common202();
  variable_set[994] = 1;
  return result;
}

static ExcelValue model_d67() {
  static ExcelValue result;
  if(variable_set[995] == 1) { return result;}
  result = _common203();
  variable_set[995] = 1;
  return result;
}

static ExcelValue model_e67() {
  static ExcelValue result;
  if(variable_set[996] == 1) { return result;}
  result = _common204();
  variable_set[996] = 1;
  return result;
}

static ExcelValue model_f67() {
  static ExcelValue result;
  if(variable_set[997] == 1) { return result;}
  result = _common205();
  variable_set[997] = 1;
  return result;
}

static ExcelValue model_g67() {
  static ExcelValue result;
  if(variable_set[998] == 1) { return result;}
  result = _common206();
  variable_set[998] = 1;
  return result;
}

static ExcelValue model_h67() {
  static ExcelValue result;
  if(variable_set[999] == 1) { return result;}
  result = _common207();
  variable_set[999] = 1;
  return result;
}

static ExcelValue model_i67() {
  static ExcelValue result;
  if(variable_set[1000] == 1) { return result;}
  result = _common208();
  variable_set[1000] = 1;
  return result;
}

static ExcelValue model_j67() {
  static ExcelValue result;
  if(variable_set[1001] == 1) { return result;}
  result = _common209();
  variable_set[1001] = 1;
  return result;
}

static ExcelValue model_k67() {
  static ExcelValue result;
  if(variable_set[1002] == 1) { return result;}
  result = _common210();
  variable_set[1002] = 1;
  return result;
}

static ExcelValue model_l67() {
  static ExcelValue result;
  if(variable_set[1003] == 1) { return result;}
  result = _common211();
  variable_set[1003] = 1;
  return result;
}

static ExcelValue model_m67() {
  static ExcelValue result;
  if(variable_set[1004] == 1) { return result;}
  result = _common212();
  variable_set[1004] = 1;
  return result;
}

static ExcelValue model_n67() {
  static ExcelValue result;
  if(variable_set[1005] == 1) { return result;}
  result = _common213();
  variable_set[1005] = 1;
  return result;
}

static ExcelValue model_o67() {
  static ExcelValue result;
  if(variable_set[1006] == 1) { return result;}
  result = _common214();
  variable_set[1006] = 1;
  return result;
}

static ExcelValue model_p67() {
  static ExcelValue result;
  if(variable_set[1007] == 1) { return result;}
  result = _common215();
  variable_set[1007] = 1;
  return result;
}

static ExcelValue model_q67() {
  static ExcelValue result;
  if(variable_set[1008] == 1) { return result;}
  result = _common216();
  variable_set[1008] = 1;
  return result;
}

static ExcelValue model_r67() {
  static ExcelValue result;
  if(variable_set[1009] == 1) { return result;}
  result = _common217();
  variable_set[1009] = 1;
  return result;
}

static ExcelValue model_s67() {
  static ExcelValue result;
  if(variable_set[1010] == 1) { return result;}
  result = _common218();
  variable_set[1010] = 1;
  return result;
}

static ExcelValue model_t67() {
  static ExcelValue result;
  if(variable_set[1011] == 1) { return result;}
  result = _common219();
  variable_set[1011] = 1;
  return result;
}

static ExcelValue model_u67() {
  static ExcelValue result;
  if(variable_set[1012] == 1) { return result;}
  result = _common220();
  variable_set[1012] = 1;
  return result;
}

static ExcelValue model_v67() {
  static ExcelValue result;
  if(variable_set[1013] == 1) { return result;}
  result = _common221();
  variable_set[1013] = 1;
  return result;
}

static ExcelValue model_w67() {
  static ExcelValue result;
  if(variable_set[1014] == 1) { return result;}
  result = _common222();
  variable_set[1014] = 1;
  return result;
}

static ExcelValue model_x67() {
  static ExcelValue result;
  if(variable_set[1015] == 1) { return result;}
  result = _common223();
  variable_set[1015] = 1;
  return result;
}

static ExcelValue model_y67() {
  static ExcelValue result;
  if(variable_set[1016] == 1) { return result;}
  result = _common224();
  variable_set[1016] = 1;
  return result;
}

static ExcelValue model_z67() {
  static ExcelValue result;
  if(variable_set[1017] == 1) { return result;}
  result = _common225();
  variable_set[1017] = 1;
  return result;
}

static ExcelValue model_aa67() {
  static ExcelValue result;
  if(variable_set[1018] == 1) { return result;}
  result = _common226();
  variable_set[1018] = 1;
  return result;
}

static ExcelValue model_ab67() {
  static ExcelValue result;
  if(variable_set[1019] == 1) { return result;}
  result = _common227();
  variable_set[1019] = 1;
  return result;
}

static ExcelValue model_ac67() {
  static ExcelValue result;
  if(variable_set[1020] == 1) { return result;}
  result = _common228();
  variable_set[1020] = 1;
  return result;
}

static ExcelValue model_ad67() {
  static ExcelValue result;
  if(variable_set[1021] == 1) { return result;}
  result = _common229();
  variable_set[1021] = 1;
  return result;
}

static ExcelValue model_ae67() {
  static ExcelValue result;
  if(variable_set[1022] == 1) { return result;}
  result = _common230();
  variable_set[1022] = 1;
  return result;
}

static ExcelValue model_af67() {
  static ExcelValue result;
  if(variable_set[1023] == 1) { return result;}
  result = _common231();
  variable_set[1023] = 1;
  return result;
}

static ExcelValue model_ag67() {
  static ExcelValue result;
  if(variable_set[1024] == 1) { return result;}
  result = _common232();
  variable_set[1024] = 1;
  return result;
}

static ExcelValue model_ah67() {
  static ExcelValue result;
  if(variable_set[1025] == 1) { return result;}
  result = _common233();
  variable_set[1025] = 1;
  return result;
}

static ExcelValue model_ai67() {
  static ExcelValue result;
  if(variable_set[1026] == 1) { return result;}
  result = _common234();
  variable_set[1026] = 1;
  return result;
}

static ExcelValue model_aj67() {
  static ExcelValue result;
  if(variable_set[1027] == 1) { return result;}
  result = _common235();
  variable_set[1027] = 1;
  return result;
}

static ExcelValue model_ak67() {
  static ExcelValue result;
  if(variable_set[1028] == 1) { return result;}
  result = _common236();
  variable_set[1028] = 1;
  return result;
}

static ExcelValue model_al67() {
  static ExcelValue result;
  if(variable_set[1029] == 1) { return result;}
  result = _common237();
  variable_set[1029] = 1;
  return result;
}

static ExcelValue model_am67() {
  static ExcelValue result;
  if(variable_set[1030] == 1) { return result;}
  result = _common238();
  variable_set[1030] = 1;
  return result;
}

static ExcelValue model_an67() {
  static ExcelValue result;
  if(variable_set[1031] == 1) { return result;}
  result = _common239();
  variable_set[1031] = 1;
  return result;
}

static ExcelValue model_b68() {
  static ExcelValue result;
  if(variable_set[1032] == 1) { return result;}
  result = subtract(C40,C10);
  variable_set[1032] = 1;
  return result;
}

static ExcelValue model_c68() {
  static ExcelValue result;
  if(variable_set[1033] == 1) { return result;}
  result = subtract(model_c64(),C25);
  variable_set[1033] = 1;
  return result;
}

static ExcelValue model_d68() {
  static ExcelValue result;
  if(variable_set[1034] == 1) { return result;}
  result = subtract(model_d64(),model_d48());
  variable_set[1034] = 1;
  return result;
}

static ExcelValue model_e68() {
  static ExcelValue result;
  if(variable_set[1035] == 1) { return result;}
  result = subtract(model_e64(),model_e48());
  variable_set[1035] = 1;
  return result;
}

static ExcelValue model_f68() {
  static ExcelValue result;
  if(variable_set[1036] == 1) { return result;}
  result = subtract(model_f64(),model_f48());
  variable_set[1036] = 1;
  return result;
}

static ExcelValue model_g68() {
  static ExcelValue result;
  if(variable_set[1037] == 1) { return result;}
  result = subtract(model_g64(),model_g48());
  variable_set[1037] = 1;
  return result;
}

static ExcelValue model_h68() {
  static ExcelValue result;
  if(variable_set[1038] == 1) { return result;}
  result = subtract(model_h64(),model_h48());
  variable_set[1038] = 1;
  return result;
}

static ExcelValue model_i68() {
  static ExcelValue result;
  if(variable_set[1039] == 1) { return result;}
  result = subtract(model_i64(),model_i48());
  variable_set[1039] = 1;
  return result;
}

static ExcelValue model_j68() {
  static ExcelValue result;
  if(variable_set[1040] == 1) { return result;}
  result = subtract(model_j64(),model_j48());
  variable_set[1040] = 1;
  return result;
}

static ExcelValue model_k68() {
  static ExcelValue result;
  if(variable_set[1041] == 1) { return result;}
  result = subtract(model_k64(),model_k48());
  variable_set[1041] = 1;
  return result;
}

static ExcelValue model_l68() {
  static ExcelValue result;
  if(variable_set[1042] == 1) { return result;}
  result = subtract(model_l64(),model_l48());
  variable_set[1042] = 1;
  return result;
}

static ExcelValue model_m68() {
  static ExcelValue result;
  if(variable_set[1043] == 1) { return result;}
  result = subtract(model_m64(),model_m48());
  variable_set[1043] = 1;
  return result;
}

static ExcelValue model_n68() {
  static ExcelValue result;
  if(variable_set[1044] == 1) { return result;}
  result = subtract(model_n64(),model_n48());
  variable_set[1044] = 1;
  return result;
}

static ExcelValue model_o68() {
  static ExcelValue result;
  if(variable_set[1045] == 1) { return result;}
  result = subtract(model_o64(),model_o48());
  variable_set[1045] = 1;
  return result;
}

static ExcelValue model_p68() {
  static ExcelValue result;
  if(variable_set[1046] == 1) { return result;}
  result = subtract(model_p64(),model_p48());
  variable_set[1046] = 1;
  return result;
}

static ExcelValue model_q68() {
  static ExcelValue result;
  if(variable_set[1047] == 1) { return result;}
  result = subtract(model_q64(),model_q48());
  variable_set[1047] = 1;
  return result;
}

static ExcelValue model_r68() {
  static ExcelValue result;
  if(variable_set[1048] == 1) { return result;}
  result = subtract(model_r64(),model_r48());
  variable_set[1048] = 1;
  return result;
}

static ExcelValue model_s68() {
  static ExcelValue result;
  if(variable_set[1049] == 1) { return result;}
  result = subtract(model_s64(),model_s48());
  variable_set[1049] = 1;
  return result;
}

static ExcelValue model_t68() {
  static ExcelValue result;
  if(variable_set[1050] == 1) { return result;}
  result = subtract(model_t64(),model_t48());
  variable_set[1050] = 1;
  return result;
}

static ExcelValue model_u68() {
  static ExcelValue result;
  if(variable_set[1051] == 1) { return result;}
  result = subtract(model_u64(),model_u48());
  variable_set[1051] = 1;
  return result;
}

static ExcelValue model_v68() {
  static ExcelValue result;
  if(variable_set[1052] == 1) { return result;}
  result = subtract(model_v64(),model_v48());
  variable_set[1052] = 1;
  return result;
}

static ExcelValue model_w68() {
  static ExcelValue result;
  if(variable_set[1053] == 1) { return result;}
  result = subtract(model_w64(),model_w48());
  variable_set[1053] = 1;
  return result;
}

static ExcelValue model_x68() {
  static ExcelValue result;
  if(variable_set[1054] == 1) { return result;}
  result = subtract(model_x64(),model_x48());
  variable_set[1054] = 1;
  return result;
}

static ExcelValue model_y68() {
  static ExcelValue result;
  if(variable_set[1055] == 1) { return result;}
  result = subtract(model_y64(),model_y48());
  variable_set[1055] = 1;
  return result;
}

static ExcelValue model_z68() {
  static ExcelValue result;
  if(variable_set[1056] == 1) { return result;}
  result = subtract(model_z64(),model_z48());
  variable_set[1056] = 1;
  return result;
}

static ExcelValue model_aa68() {
  static ExcelValue result;
  if(variable_set[1057] == 1) { return result;}
  result = subtract(model_aa64(),model_aa48());
  variable_set[1057] = 1;
  return result;
}

static ExcelValue model_ab68() {
  static ExcelValue result;
  if(variable_set[1058] == 1) { return result;}
  result = subtract(model_ab64(),model_ab48());
  variable_set[1058] = 1;
  return result;
}

static ExcelValue model_ac68() {
  static ExcelValue result;
  if(variable_set[1059] == 1) { return result;}
  result = subtract(model_ac64(),model_ac48());
  variable_set[1059] = 1;
  return result;
}

static ExcelValue model_ad68() {
  static ExcelValue result;
  if(variable_set[1060] == 1) { return result;}
  result = subtract(model_ad64(),model_ad48());
  variable_set[1060] = 1;
  return result;
}

static ExcelValue model_ae68() {
  static ExcelValue result;
  if(variable_set[1061] == 1) { return result;}
  result = subtract(model_ae64(),model_ae48());
  variable_set[1061] = 1;
  return result;
}

static ExcelValue model_af68() {
  static ExcelValue result;
  if(variable_set[1062] == 1) { return result;}
  result = subtract(model_af64(),model_af48());
  variable_set[1062] = 1;
  return result;
}

static ExcelValue model_ag68() {
  static ExcelValue result;
  if(variable_set[1063] == 1) { return result;}
  result = subtract(model_ag64(),model_ag48());
  variable_set[1063] = 1;
  return result;
}

static ExcelValue model_ah68() {
  static ExcelValue result;
  if(variable_set[1064] == 1) { return result;}
  result = subtract(model_ah64(),model_ah48());
  variable_set[1064] = 1;
  return result;
}

static ExcelValue model_ai68() {
  static ExcelValue result;
  if(variable_set[1065] == 1) { return result;}
  result = subtract(model_ai64(),model_ai48());
  variable_set[1065] = 1;
  return result;
}

static ExcelValue model_aj68() {
  static ExcelValue result;
  if(variable_set[1066] == 1) { return result;}
  result = subtract(model_aj64(),model_aj48());
  variable_set[1066] = 1;
  return result;
}

static ExcelValue model_ak68() {
  static ExcelValue result;
  if(variable_set[1067] == 1) { return result;}
  result = subtract(model_ak64(),model_ak48());
  variable_set[1067] = 1;
  return result;
}

static ExcelValue model_al68() {
  static ExcelValue result;
  if(variable_set[1068] == 1) { return result;}
  result = subtract(model_al64(),model_al48());
  variable_set[1068] = 1;
  return result;
}

static ExcelValue model_am68() {
  static ExcelValue result;
  if(variable_set[1069] == 1) { return result;}
  result = subtract(model_am64(),model_am48());
  variable_set[1069] = 1;
  return result;
}

static ExcelValue model_an68() {
  static ExcelValue result;
  if(variable_set[1070] == 1) { return result;}
  result = subtract(model_an64(),model_an48());
  variable_set[1070] = 1;
  return result;
}

static ExcelValue model_b72() {
  static ExcelValue result;
  if(variable_set[1071] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common201(),divide(model_b67(),C23))),add(model_b68(),model_b67()));
  variable_set[1071] = 1;
  return result;
}

static ExcelValue model_c72() {
  static ExcelValue result;
  if(variable_set[1072] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common202(),divide(model_c67(),C23))),add(model_c68(),model_c67()));
  variable_set[1072] = 1;
  return result;
}

static ExcelValue model_d72() {
  static ExcelValue result;
  if(variable_set[1073] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common203(),divide(model_d67(),C23))),add(model_d68(),model_d67()));
  variable_set[1073] = 1;
  return result;
}

static ExcelValue model_e72() {
  static ExcelValue result;
  if(variable_set[1074] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common204(),divide(model_e67(),C23))),add(model_e68(),model_e67()));
  variable_set[1074] = 1;
  return result;
}

static ExcelValue model_f72() {
  static ExcelValue result;
  if(variable_set[1075] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common205(),divide(model_f67(),C23))),add(model_f68(),model_f67()));
  variable_set[1075] = 1;
  return result;
}

static ExcelValue model_g72() {
  static ExcelValue result;
  if(variable_set[1076] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common206(),divide(model_g67(),C23))),add(model_g68(),model_g67()));
  variable_set[1076] = 1;
  return result;
}

static ExcelValue model_h72() {
  static ExcelValue result;
  if(variable_set[1077] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common207(),divide(model_h67(),C23))),add(model_h68(),model_h67()));
  variable_set[1077] = 1;
  return result;
}

static ExcelValue model_i72() {
  static ExcelValue result;
  if(variable_set[1078] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common208(),divide(model_i67(),C23))),add(model_i68(),model_i67()));
  variable_set[1078] = 1;
  return result;
}

static ExcelValue model_j72() {
  static ExcelValue result;
  if(variable_set[1079] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common209(),divide(model_j67(),C23))),add(model_j68(),model_j67()));
  variable_set[1079] = 1;
  return result;
}

static ExcelValue model_k72() {
  static ExcelValue result;
  if(variable_set[1080] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common210(),divide(model_k67(),C23))),add(model_k68(),model_k67()));
  variable_set[1080] = 1;
  return result;
}

static ExcelValue model_l72() {
  static ExcelValue result;
  if(variable_set[1081] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common211(),divide(model_l67(),C23))),add(model_l68(),model_l67()));
  variable_set[1081] = 1;
  return result;
}

static ExcelValue model_m72() {
  static ExcelValue result;
  if(variable_set[1082] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common212(),divide(model_m67(),C23))),add(model_m68(),model_m67()));
  variable_set[1082] = 1;
  return result;
}

static ExcelValue model_n72() {
  static ExcelValue result;
  if(variable_set[1083] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common213(),divide(model_n67(),C23))),add(model_n68(),model_n67()));
  variable_set[1083] = 1;
  return result;
}

static ExcelValue model_o72() {
  static ExcelValue result;
  if(variable_set[1084] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common214(),divide(model_o67(),C23))),add(model_o68(),model_o67()));
  variable_set[1084] = 1;
  return result;
}

static ExcelValue model_p72() {
  static ExcelValue result;
  if(variable_set[1085] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common215(),divide(model_p67(),C23))),add(model_p68(),model_p67()));
  variable_set[1085] = 1;
  return result;
}

static ExcelValue model_q72() {
  static ExcelValue result;
  if(variable_set[1086] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common216(),divide(model_q67(),C23))),add(model_q68(),model_q67()));
  variable_set[1086] = 1;
  return result;
}

static ExcelValue model_r72() {
  static ExcelValue result;
  if(variable_set[1087] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common217(),divide(model_r67(),C23))),add(model_r68(),model_r67()));
  variable_set[1087] = 1;
  return result;
}

static ExcelValue model_s72() {
  static ExcelValue result;
  if(variable_set[1088] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common218(),divide(model_s67(),C23))),add(model_s68(),model_s67()));
  variable_set[1088] = 1;
  return result;
}

static ExcelValue model_t72() {
  static ExcelValue result;
  if(variable_set[1089] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common219(),divide(model_t67(),C23))),add(model_t68(),model_t67()));
  variable_set[1089] = 1;
  return result;
}

static ExcelValue model_u72() {
  static ExcelValue result;
  if(variable_set[1090] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common220(),divide(model_u67(),C23))),add(model_u68(),model_u67()));
  variable_set[1090] = 1;
  return result;
}

static ExcelValue model_v72() {
  static ExcelValue result;
  if(variable_set[1091] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common221(),divide(model_v67(),C23))),add(model_v68(),model_v67()));
  variable_set[1091] = 1;
  return result;
}

static ExcelValue model_w72() {
  static ExcelValue result;
  if(variable_set[1092] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common222(),divide(model_w67(),C23))),add(model_w68(),model_w67()));
  variable_set[1092] = 1;
  return result;
}

static ExcelValue model_x72() {
  static ExcelValue result;
  if(variable_set[1093] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common223(),divide(model_x67(),C23))),add(model_x68(),model_x67()));
  variable_set[1093] = 1;
  return result;
}

static ExcelValue model_y72() {
  static ExcelValue result;
  if(variable_set[1094] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common224(),divide(model_y67(),C23))),add(model_y68(),model_y67()));
  variable_set[1094] = 1;
  return result;
}

static ExcelValue model_z72() {
  static ExcelValue result;
  if(variable_set[1095] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common225(),divide(model_z67(),C23))),add(model_z68(),model_z67()));
  variable_set[1095] = 1;
  return result;
}

static ExcelValue model_aa72() {
  static ExcelValue result;
  if(variable_set[1096] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common226(),divide(model_aa67(),C23))),add(model_aa68(),model_aa67()));
  variable_set[1096] = 1;
  return result;
}

static ExcelValue model_ab72() {
  static ExcelValue result;
  if(variable_set[1097] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common227(),divide(model_ab67(),C23))),add(model_ab68(),model_ab67()));
  variable_set[1097] = 1;
  return result;
}

static ExcelValue model_ac72() {
  static ExcelValue result;
  if(variable_set[1098] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common228(),divide(model_ac67(),C23))),add(model_ac68(),model_ac67()));
  variable_set[1098] = 1;
  return result;
}

static ExcelValue model_ad72() {
  static ExcelValue result;
  if(variable_set[1099] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common229(),divide(model_ad67(),C23))),add(model_ad68(),model_ad67()));
  variable_set[1099] = 1;
  return result;
}

static ExcelValue model_ae72() {
  static ExcelValue result;
  if(variable_set[1100] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common230(),divide(model_ae67(),C23))),add(model_ae68(),model_ae67()));
  variable_set[1100] = 1;
  return result;
}

static ExcelValue model_af72() {
  static ExcelValue result;
  if(variable_set[1101] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common231(),divide(model_af67(),C23))),add(model_af68(),model_af67()));
  variable_set[1101] = 1;
  return result;
}

static ExcelValue model_ag72() {
  static ExcelValue result;
  if(variable_set[1102] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common232(),divide(model_ag67(),C23))),add(model_ag68(),model_ag67()));
  variable_set[1102] = 1;
  return result;
}

static ExcelValue model_ah72() {
  static ExcelValue result;
  if(variable_set[1103] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common233(),divide(model_ah67(),C23))),add(model_ah68(),model_ah67()));
  variable_set[1103] = 1;
  return result;
}

static ExcelValue model_ai72() {
  static ExcelValue result;
  if(variable_set[1104] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common234(),divide(model_ai67(),C23))),add(model_ai68(),model_ai67()));
  variable_set[1104] = 1;
  return result;
}

static ExcelValue model_aj72() {
  static ExcelValue result;
  if(variable_set[1105] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common235(),divide(model_aj67(),C23))),add(model_aj68(),model_aj67()));
  variable_set[1105] = 1;
  return result;
}

static ExcelValue model_ak72() {
  static ExcelValue result;
  if(variable_set[1106] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common236(),divide(model_ak67(),C23))),add(model_ak68(),model_ak67()));
  variable_set[1106] = 1;
  return result;
}

static ExcelValue model_al72() {
  static ExcelValue result;
  if(variable_set[1107] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common237(),divide(model_al67(),C23))),add(model_al68(),model_al67()));
  variable_set[1107] = 1;
  return result;
}

static ExcelValue model_am72() {
  static ExcelValue result;
  if(variable_set[1108] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common238(),divide(model_am67(),C23))),add(model_am68(),model_am67()));
  variable_set[1108] = 1;
  return result;
}

static ExcelValue model_an72() {
  static ExcelValue result;
  if(variable_set[1109] == 1) { return result;}
  result = divide(multiply(C23,subtract(_common239(),divide(model_an67(),C23))),add(model_an68(),model_an67()));
  variable_set[1109] = 1;
  return result;
}

static ExcelValue model_k74() {
  static ExcelValue result;
  if(variable_set[1110] == 1) { return result;}
  result = add(model_j49(),_common135());
  variable_set[1110] = 1;
  return result;
}

static ExcelValue model_l74() {
  static ExcelValue result;
  if(variable_set[1111] == 1) { return result;}
  result = add(model_k74(),_common136());
  variable_set[1111] = 1;
  return result;
}

static ExcelValue model_m74() {
  static ExcelValue result;
  if(variable_set[1112] == 1) { return result;}
  result = add(model_l74(),_common137());
  variable_set[1112] = 1;
  return result;
}

static ExcelValue model_n74() {
  static ExcelValue result;
  if(variable_set[1113] == 1) { return result;}
  result = add(model_m74(),_common138());
  variable_set[1113] = 1;
  return result;
}

static ExcelValue model_o74() {
  static ExcelValue result;
  if(variable_set[1114] == 1) { return result;}
  result = add(model_n74(),_common139());
  variable_set[1114] = 1;
  return result;
}

static ExcelValue model_p74() {
  static ExcelValue result;
  if(variable_set[1115] == 1) { return result;}
  result = add(model_o74(),_common140());
  variable_set[1115] = 1;
  return result;
}

static ExcelValue model_q74() {
  static ExcelValue result;
  if(variable_set[1116] == 1) { return result;}
  result = add(model_p74(),_common141());
  variable_set[1116] = 1;
  return result;
}

static ExcelValue model_r74() {
  static ExcelValue result;
  if(variable_set[1117] == 1) { return result;}
  result = add(model_q74(),_common142());
  variable_set[1117] = 1;
  return result;
}

static ExcelValue model_s74() {
  static ExcelValue result;
  if(variable_set[1118] == 1) { return result;}
  result = add(model_r74(),_common143());
  variable_set[1118] = 1;
  return result;
}

static ExcelValue model_t74() {
  static ExcelValue result;
  if(variable_set[1119] == 1) { return result;}
  result = add(model_s74(),_common144());
  variable_set[1119] = 1;
  return result;
}

static ExcelValue model_u74() {
  static ExcelValue result;
  if(variable_set[1120] == 1) { return result;}
  result = add(model_t74(),_common145());
  variable_set[1120] = 1;
  return result;
}

static ExcelValue model_v74() {
  static ExcelValue result;
  if(variable_set[1121] == 1) { return result;}
  result = add(model_u74(),_common146());
  variable_set[1121] = 1;
  return result;
}

static ExcelValue model_w74() {
  static ExcelValue result;
  if(variable_set[1122] == 1) { return result;}
  result = add(model_v74(),_common147());
  variable_set[1122] = 1;
  return result;
}

static ExcelValue model_x74() {
  static ExcelValue result;
  if(variable_set[1123] == 1) { return result;}
  result = add(model_w74(),_common148());
  variable_set[1123] = 1;
  return result;
}

static ExcelValue model_y74() {
  static ExcelValue result;
  if(variable_set[1124] == 1) { return result;}
  result = add(model_x74(),_common149());
  variable_set[1124] = 1;
  return result;
}

static ExcelValue model_z74() {
  static ExcelValue result;
  if(variable_set[1125] == 1) { return result;}
  result = add(model_y74(),_common150());
  variable_set[1125] = 1;
  return result;
}

static ExcelValue model_aa74() {
  static ExcelValue result;
  if(variable_set[1126] == 1) { return result;}
  result = add(model_z74(),_common151());
  variable_set[1126] = 1;
  return result;
}

static ExcelValue model_ab74() {
  static ExcelValue result;
  if(variable_set[1127] == 1) { return result;}
  result = add(model_aa74(),_common152());
  variable_set[1127] = 1;
  return result;
}

static ExcelValue model_ac74() {
  static ExcelValue result;
  if(variable_set[1128] == 1) { return result;}
  result = add(model_ab74(),_common153());
  variable_set[1128] = 1;
  return result;
}

static ExcelValue model_ad74() {
  static ExcelValue result;
  if(variable_set[1129] == 1) { return result;}
  result = add(model_ac74(),_common154());
  variable_set[1129] = 1;
  return result;
}

static ExcelValue model_ae74() {
  static ExcelValue result;
  if(variable_set[1130] == 1) { return result;}
  result = add(model_ad74(),_common155());
  variable_set[1130] = 1;
  return result;
}

static ExcelValue model_af74() {
  static ExcelValue result;
  if(variable_set[1131] == 1) { return result;}
  result = add(model_ae74(),_common156());
  variable_set[1131] = 1;
  return result;
}

static ExcelValue model_ag74() {
  static ExcelValue result;
  if(variable_set[1132] == 1) { return result;}
  result = add(model_af74(),_common157());
  variable_set[1132] = 1;
  return result;
}

static ExcelValue model_ah74() {
  static ExcelValue result;
  if(variable_set[1133] == 1) { return result;}
  result = add(model_ag74(),_common158());
  variable_set[1133] = 1;
  return result;
}

static ExcelValue model_ai74() {
  static ExcelValue result;
  if(variable_set[1134] == 1) { return result;}
  result = add(model_ah74(),_common159());
  variable_set[1134] = 1;
  return result;
}

static ExcelValue model_aj74() {
  static ExcelValue result;
  if(variable_set[1135] == 1) { return result;}
  result = add(model_ai74(),_common160());
  variable_set[1135] = 1;
  return result;
}

static ExcelValue model_ak74() {
  static ExcelValue result;
  if(variable_set[1136] == 1) { return result;}
  result = add(model_aj74(),_common161());
  variable_set[1136] = 1;
  return result;
}

static ExcelValue model_al74() {
  static ExcelValue result;
  if(variable_set[1137] == 1) { return result;}
  result = add(model_ak74(),_common162());
  variable_set[1137] = 1;
  return result;
}

static ExcelValue model_am74() {
  static ExcelValue result;
  if(variable_set[1138] == 1) { return result;}
  result = add(model_al74(),_common163());
  variable_set[1138] = 1;
  return result;
}

static ExcelValue model_an74() {
  static ExcelValue result;
  if(variable_set[1139] == 1) { return result;}
  result = add(model_am74(),_common164());
  variable_set[1139] = 1;
  return result;
}

static ExcelValue model_b75() {
  static ExcelValue result;
  if(variable_set[1140] == 1) { return result;}
  ExcelValue array0[] = {C17,C41};
  result = min(2, array0);
  variable_set[1140] = 1;
  return result;
}

static ExcelValue model_c75() {
  static ExcelValue result;
  if(variable_set[1141] == 1) { return result;}
  ExcelValue array0[] = {model_c49(),model_c63()};
  result = min(2, array0);
  variable_set[1141] = 1;
  return result;
}

static ExcelValue model_d75() {
  static ExcelValue result;
  if(variable_set[1142] == 1) { return result;}
  ExcelValue array0[] = {model_d49(),model_d63()};
  result = min(2, array0);
  variable_set[1142] = 1;
  return result;
}

static ExcelValue model_e75() {
  static ExcelValue result;
  if(variable_set[1143] == 1) { return result;}
  ExcelValue array0[] = {model_e49(),model_e63()};
  result = min(2, array0);
  variable_set[1143] = 1;
  return result;
}

static ExcelValue model_f75() {
  static ExcelValue result;
  if(variable_set[1144] == 1) { return result;}
  ExcelValue array0[] = {model_f49(),model_f63()};
  result = min(2, array0);
  variable_set[1144] = 1;
  return result;
}

static ExcelValue model_g75() {
  static ExcelValue result;
  if(variable_set[1145] == 1) { return result;}
  ExcelValue array0[] = {model_g49(),model_g63()};
  result = min(2, array0);
  variable_set[1145] = 1;
  return result;
}

static ExcelValue model_h75() {
  static ExcelValue result;
  if(variable_set[1146] == 1) { return result;}
  ExcelValue array0[] = {model_h49(),model_h63()};
  result = min(2, array0);
  variable_set[1146] = 1;
  return result;
}

static ExcelValue model_i75() {
  static ExcelValue result;
  if(variable_set[1147] == 1) { return result;}
  ExcelValue array0[] = {model_i49(),model_i63()};
  result = min(2, array0);
  variable_set[1147] = 1;
  return result;
}

static ExcelValue model_j75() {
  static ExcelValue result;
  if(variable_set[1148] == 1) { return result;}
  ExcelValue array0[] = {model_j49(),model_j63()};
  result = min(2, array0);
  variable_set[1148] = 1;
  return result;
}

static ExcelValue model_k75() {
  static ExcelValue result;
  if(variable_set[1149] == 1) { return result;}
  ExcelValue array0[] = {model_k74(),model_k63()};
  result = min(2, array0);
  variable_set[1149] = 1;
  return result;
}

static ExcelValue model_l75() {
  static ExcelValue result;
  if(variable_set[1150] == 1) { return result;}
  ExcelValue array0[] = {model_l74(),model_l63()};
  result = min(2, array0);
  variable_set[1150] = 1;
  return result;
}

static ExcelValue model_m75() {
  static ExcelValue result;
  if(variable_set[1151] == 1) { return result;}
  ExcelValue array0[] = {model_m74(),model_m63()};
  result = min(2, array0);
  variable_set[1151] = 1;
  return result;
}

static ExcelValue model_n75() {
  static ExcelValue result;
  if(variable_set[1152] == 1) { return result;}
  ExcelValue array0[] = {model_n74(),model_n63()};
  result = min(2, array0);
  variable_set[1152] = 1;
  return result;
}

static ExcelValue model_o75() {
  static ExcelValue result;
  if(variable_set[1153] == 1) { return result;}
  ExcelValue array0[] = {model_o74(),model_o63()};
  result = min(2, array0);
  variable_set[1153] = 1;
  return result;
}

static ExcelValue model_p75() {
  static ExcelValue result;
  if(variable_set[1154] == 1) { return result;}
  ExcelValue array0[] = {model_p74(),model_p63()};
  result = min(2, array0);
  variable_set[1154] = 1;
  return result;
}

static ExcelValue model_q75() {
  static ExcelValue result;
  if(variable_set[1155] == 1) { return result;}
  ExcelValue array0[] = {model_q74(),model_q63()};
  result = min(2, array0);
  variable_set[1155] = 1;
  return result;
}

static ExcelValue model_r75() {
  static ExcelValue result;
  if(variable_set[1156] == 1) { return result;}
  ExcelValue array0[] = {model_r74(),model_r63()};
  result = min(2, array0);
  variable_set[1156] = 1;
  return result;
}

static ExcelValue model_s75() {
  static ExcelValue result;
  if(variable_set[1157] == 1) { return result;}
  ExcelValue array0[] = {model_s74(),model_s63()};
  result = min(2, array0);
  variable_set[1157] = 1;
  return result;
}

static ExcelValue model_t75() {
  static ExcelValue result;
  if(variable_set[1158] == 1) { return result;}
  ExcelValue array0[] = {model_t74(),model_t63()};
  result = min(2, array0);
  variable_set[1158] = 1;
  return result;
}

static ExcelValue model_u75() {
  static ExcelValue result;
  if(variable_set[1159] == 1) { return result;}
  ExcelValue array0[] = {model_u74(),model_u63()};
  result = min(2, array0);
  variable_set[1159] = 1;
  return result;
}

static ExcelValue model_v75() {
  static ExcelValue result;
  if(variable_set[1160] == 1) { return result;}
  ExcelValue array0[] = {model_v74(),model_v63()};
  result = min(2, array0);
  variable_set[1160] = 1;
  return result;
}

static ExcelValue model_w75() {
  static ExcelValue result;
  if(variable_set[1161] == 1) { return result;}
  ExcelValue array0[] = {model_w74(),model_w63()};
  result = min(2, array0);
  variable_set[1161] = 1;
  return result;
}

static ExcelValue model_x75() {
  static ExcelValue result;
  if(variable_set[1162] == 1) { return result;}
  ExcelValue array0[] = {model_x74(),model_x63()};
  result = min(2, array0);
  variable_set[1162] = 1;
  return result;
}

static ExcelValue model_y75() {
  static ExcelValue result;
  if(variable_set[1163] == 1) { return result;}
  ExcelValue array0[] = {model_y74(),model_y63()};
  result = min(2, array0);
  variable_set[1163] = 1;
  return result;
}

static ExcelValue model_z75() {
  static ExcelValue result;
  if(variable_set[1164] == 1) { return result;}
  ExcelValue array0[] = {model_z74(),model_z63()};
  result = min(2, array0);
  variable_set[1164] = 1;
  return result;
}

static ExcelValue model_aa75() {
  static ExcelValue result;
  if(variable_set[1165] == 1) { return result;}
  ExcelValue array0[] = {model_aa74(),model_aa63()};
  result = min(2, array0);
  variable_set[1165] = 1;
  return result;
}

static ExcelValue model_ab75() {
  static ExcelValue result;
  if(variable_set[1166] == 1) { return result;}
  ExcelValue array0[] = {model_ab74(),model_ab63()};
  result = min(2, array0);
  variable_set[1166] = 1;
  return result;
}

static ExcelValue model_ac75() {
  static ExcelValue result;
  if(variable_set[1167] == 1) { return result;}
  ExcelValue array0[] = {model_ac74(),model_ac63()};
  result = min(2, array0);
  variable_set[1167] = 1;
  return result;
}

static ExcelValue model_ad75() {
  static ExcelValue result;
  if(variable_set[1168] == 1) { return result;}
  ExcelValue array0[] = {model_ad74(),model_ad63()};
  result = min(2, array0);
  variable_set[1168] = 1;
  return result;
}

static ExcelValue model_ae75() {
  static ExcelValue result;
  if(variable_set[1169] == 1) { return result;}
  ExcelValue array0[] = {model_ae74(),model_ae63()};
  result = min(2, array0);
  variable_set[1169] = 1;
  return result;
}

static ExcelValue model_af75() {
  static ExcelValue result;
  if(variable_set[1170] == 1) { return result;}
  ExcelValue array0[] = {model_af74(),model_af63()};
  result = min(2, array0);
  variable_set[1170] = 1;
  return result;
}

static ExcelValue model_ag75() {
  static ExcelValue result;
  if(variable_set[1171] == 1) { return result;}
  ExcelValue array0[] = {model_ag74(),model_ag63()};
  result = min(2, array0);
  variable_set[1171] = 1;
  return result;
}

static ExcelValue model_ah75() {
  static ExcelValue result;
  if(variable_set[1172] == 1) { return result;}
  ExcelValue array0[] = {model_ah74(),model_ah63()};
  result = min(2, array0);
  variable_set[1172] = 1;
  return result;
}

static ExcelValue model_ai75() {
  static ExcelValue result;
  if(variable_set[1173] == 1) { return result;}
  ExcelValue array0[] = {model_ai74(),model_ai63()};
  result = min(2, array0);
  variable_set[1173] = 1;
  return result;
}

static ExcelValue model_aj75() {
  static ExcelValue result;
  if(variable_set[1174] == 1) { return result;}
  ExcelValue array0[] = {model_aj74(),model_aj63()};
  result = min(2, array0);
  variable_set[1174] = 1;
  return result;
}

static ExcelValue model_ak75() {
  static ExcelValue result;
  if(variable_set[1175] == 1) { return result;}
  ExcelValue array0[] = {model_ak74(),model_ak63()};
  result = min(2, array0);
  variable_set[1175] = 1;
  return result;
}

static ExcelValue model_al75() {
  static ExcelValue result;
  if(variable_set[1176] == 1) { return result;}
  ExcelValue array0[] = {model_al74(),model_al63()};
  result = min(2, array0);
  variable_set[1176] = 1;
  return result;
}

static ExcelValue model_am75() {
  static ExcelValue result;
  if(variable_set[1177] == 1) { return result;}
  ExcelValue array0[] = {model_am74(),model_am63()};
  result = min(2, array0);
  variable_set[1177] = 1;
  return result;
}

static ExcelValue model_an75() {
  static ExcelValue result;
  if(variable_set[1178] == 1) { return result;}
  ExcelValue array0[] = {model_an74(),model_an63()};
  result = min(2, array0);
  variable_set[1178] = 1;
  return result;
}

static ExcelValue model_b76() {
  static ExcelValue result;
  if(variable_set[1179] == 1) { return result;}
  ExcelValue array0[] = {model_b67(),_common240()};
  result = min(2, array0);
  variable_set[1179] = 1;
  return result;
}

static ExcelValue model_c76() {
  static ExcelValue result;
  if(variable_set[1180] == 1) { return result;}
  ExcelValue array0[] = {model_c67(),_common241()};
  result = min(2, array0);
  variable_set[1180] = 1;
  return result;
}

static ExcelValue model_d76() {
  static ExcelValue result;
  if(variable_set[1181] == 1) { return result;}
  ExcelValue array0[] = {model_d67(),_common242()};
  result = min(2, array0);
  variable_set[1181] = 1;
  return result;
}

static ExcelValue model_e76() {
  static ExcelValue result;
  if(variable_set[1182] == 1) { return result;}
  ExcelValue array0[] = {model_e67(),_common243()};
  result = min(2, array0);
  variable_set[1182] = 1;
  return result;
}

static ExcelValue model_f76() {
  static ExcelValue result;
  if(variable_set[1183] == 1) { return result;}
  ExcelValue array0[] = {model_f67(),_common244()};
  result = min(2, array0);
  variable_set[1183] = 1;
  return result;
}

static ExcelValue model_g76() {
  static ExcelValue result;
  if(variable_set[1184] == 1) { return result;}
  ExcelValue array0[] = {model_g67(),_common245()};
  result = min(2, array0);
  variable_set[1184] = 1;
  return result;
}

static ExcelValue model_h76() {
  static ExcelValue result;
  if(variable_set[1185] == 1) { return result;}
  ExcelValue array0[] = {model_h67(),_common246()};
  result = min(2, array0);
  variable_set[1185] = 1;
  return result;
}

static ExcelValue model_i76() {
  static ExcelValue result;
  if(variable_set[1186] == 1) { return result;}
  ExcelValue array0[] = {model_i67(),_common247()};
  result = min(2, array0);
  variable_set[1186] = 1;
  return result;
}

static ExcelValue model_j76() {
  static ExcelValue result;
  if(variable_set[1187] == 1) { return result;}
  ExcelValue array0[] = {model_j67(),_common248()};
  result = min(2, array0);
  variable_set[1187] = 1;
  return result;
}

static ExcelValue model_k76() {
  static ExcelValue result;
  if(variable_set[1188] == 1) { return result;}
  ExcelValue array0[] = {model_k67(),_common249()};
  result = min(2, array0);
  variable_set[1188] = 1;
  return result;
}

static ExcelValue model_l76() {
  static ExcelValue result;
  if(variable_set[1189] == 1) { return result;}
  ExcelValue array0[] = {model_l67(),_common250()};
  result = min(2, array0);
  variable_set[1189] = 1;
  return result;
}

static ExcelValue model_m76() {
  static ExcelValue result;
  if(variable_set[1190] == 1) { return result;}
  ExcelValue array0[] = {model_m67(),_common251()};
  result = min(2, array0);
  variable_set[1190] = 1;
  return result;
}

static ExcelValue model_n76() {
  static ExcelValue result;
  if(variable_set[1191] == 1) { return result;}
  ExcelValue array0[] = {model_n67(),_common252()};
  result = min(2, array0);
  variable_set[1191] = 1;
  return result;
}

static ExcelValue model_o76() {
  static ExcelValue result;
  if(variable_set[1192] == 1) { return result;}
  ExcelValue array0[] = {model_o67(),_common253()};
  result = min(2, array0);
  variable_set[1192] = 1;
  return result;
}

static ExcelValue model_p76() {
  static ExcelValue result;
  if(variable_set[1193] == 1) { return result;}
  ExcelValue array0[] = {model_p67(),_common254()};
  result = min(2, array0);
  variable_set[1193] = 1;
  return result;
}

static ExcelValue model_q76() {
  static ExcelValue result;
  if(variable_set[1194] == 1) { return result;}
  ExcelValue array0[] = {model_q67(),_common255()};
  result = min(2, array0);
  variable_set[1194] = 1;
  return result;
}

static ExcelValue model_r76() {
  static ExcelValue result;
  if(variable_set[1195] == 1) { return result;}
  ExcelValue array0[] = {model_r67(),_common256()};
  result = min(2, array0);
  variable_set[1195] = 1;
  return result;
}

static ExcelValue model_s76() {
  static ExcelValue result;
  if(variable_set[1196] == 1) { return result;}
  ExcelValue array0[] = {model_s67(),_common257()};
  result = min(2, array0);
  variable_set[1196] = 1;
  return result;
}

static ExcelValue model_t76() {
  static ExcelValue result;
  if(variable_set[1197] == 1) { return result;}
  ExcelValue array0[] = {model_t67(),_common258()};
  result = min(2, array0);
  variable_set[1197] = 1;
  return result;
}

static ExcelValue model_u76() {
  static ExcelValue result;
  if(variable_set[1198] == 1) { return result;}
  ExcelValue array0[] = {model_u67(),_common259()};
  result = min(2, array0);
  variable_set[1198] = 1;
  return result;
}

static ExcelValue model_v76() {
  static ExcelValue result;
  if(variable_set[1199] == 1) { return result;}
  ExcelValue array0[] = {model_v67(),_common260()};
  result = min(2, array0);
  variable_set[1199] = 1;
  return result;
}

static ExcelValue model_w76() {
  static ExcelValue result;
  if(variable_set[1200] == 1) { return result;}
  ExcelValue array0[] = {model_w67(),_common261()};
  result = min(2, array0);
  variable_set[1200] = 1;
  return result;
}

static ExcelValue model_x76() {
  static ExcelValue result;
  if(variable_set[1201] == 1) { return result;}
  ExcelValue array0[] = {model_x67(),_common262()};
  result = min(2, array0);
  variable_set[1201] = 1;
  return result;
}

static ExcelValue model_y76() {
  static ExcelValue result;
  if(variable_set[1202] == 1) { return result;}
  ExcelValue array0[] = {model_y67(),_common263()};
  result = min(2, array0);
  variable_set[1202] = 1;
  return result;
}

static ExcelValue model_z76() {
  static ExcelValue result;
  if(variable_set[1203] == 1) { return result;}
  ExcelValue array0[] = {model_z67(),_common264()};
  result = min(2, array0);
  variable_set[1203] = 1;
  return result;
}

static ExcelValue model_aa76() {
  static ExcelValue result;
  if(variable_set[1204] == 1) { return result;}
  ExcelValue array0[] = {model_aa67(),_common265()};
  result = min(2, array0);
  variable_set[1204] = 1;
  return result;
}

static ExcelValue model_ab76() {
  static ExcelValue result;
  if(variable_set[1205] == 1) { return result;}
  ExcelValue array0[] = {model_ab67(),_common266()};
  result = min(2, array0);
  variable_set[1205] = 1;
  return result;
}

static ExcelValue model_ac76() {
  static ExcelValue result;
  if(variable_set[1206] == 1) { return result;}
  ExcelValue array0[] = {model_ac67(),_common267()};
  result = min(2, array0);
  variable_set[1206] = 1;
  return result;
}

static ExcelValue model_ad76() {
  static ExcelValue result;
  if(variable_set[1207] == 1) { return result;}
  ExcelValue array0[] = {model_ad67(),_common268()};
  result = min(2, array0);
  variable_set[1207] = 1;
  return result;
}

static ExcelValue model_ae76() {
  static ExcelValue result;
  if(variable_set[1208] == 1) { return result;}
  ExcelValue array0[] = {model_ae67(),_common269()};
  result = min(2, array0);
  variable_set[1208] = 1;
  return result;
}

static ExcelValue model_af76() {
  static ExcelValue result;
  if(variable_set[1209] == 1) { return result;}
  ExcelValue array0[] = {model_af67(),_common270()};
  result = min(2, array0);
  variable_set[1209] = 1;
  return result;
}

static ExcelValue model_ag76() {
  static ExcelValue result;
  if(variable_set[1210] == 1) { return result;}
  ExcelValue array0[] = {model_ag67(),_common271()};
  result = min(2, array0);
  variable_set[1210] = 1;
  return result;
}

static ExcelValue model_ah76() {
  static ExcelValue result;
  if(variable_set[1211] == 1) { return result;}
  ExcelValue array0[] = {model_ah67(),_common272()};
  result = min(2, array0);
  variable_set[1211] = 1;
  return result;
}

static ExcelValue model_ai76() {
  static ExcelValue result;
  if(variable_set[1212] == 1) { return result;}
  ExcelValue array0[] = {model_ai67(),_common273()};
  result = min(2, array0);
  variable_set[1212] = 1;
  return result;
}

static ExcelValue model_aj76() {
  static ExcelValue result;
  if(variable_set[1213] == 1) { return result;}
  ExcelValue array0[] = {model_aj67(),_common274()};
  result = min(2, array0);
  variable_set[1213] = 1;
  return result;
}

static ExcelValue model_ak76() {
  static ExcelValue result;
  if(variable_set[1214] == 1) { return result;}
  ExcelValue array0[] = {model_ak67(),_common275()};
  result = min(2, array0);
  variable_set[1214] = 1;
  return result;
}

static ExcelValue model_al76() {
  static ExcelValue result;
  if(variable_set[1215] == 1) { return result;}
  ExcelValue array0[] = {model_al67(),_common276()};
  result = min(2, array0);
  variable_set[1215] = 1;
  return result;
}

static ExcelValue model_am76() {
  static ExcelValue result;
  if(variable_set[1216] == 1) { return result;}
  ExcelValue array0[] = {model_am67(),_common277()};
  result = min(2, array0);
  variable_set[1216] = 1;
  return result;
}

static ExcelValue model_an76() {
  static ExcelValue result;
  if(variable_set[1217] == 1) { return result;}
  ExcelValue array0[] = {model_an67(),_common278()};
  result = min(2, array0);
  variable_set[1217] = 1;
  return result;
}

static ExcelValue model_b77() {
  static ExcelValue result;
  if(variable_set[1218] == 1) { return result;}
  ExcelValue array0[] = {model_b68(),subtract(_common240(),model_b76())};
  result = min(2, array0);
  variable_set[1218] = 1;
  return result;
}

static ExcelValue model_c77() {
  static ExcelValue result;
  if(variable_set[1219] == 1) { return result;}
  ExcelValue array0[] = {model_c68(),subtract(_common241(),model_c76())};
  result = min(2, array0);
  variable_set[1219] = 1;
  return result;
}

static ExcelValue model_d77() {
  static ExcelValue result;
  if(variable_set[1220] == 1) { return result;}
  ExcelValue array0[] = {model_d68(),subtract(_common242(),model_d76())};
  result = min(2, array0);
  variable_set[1220] = 1;
  return result;
}

static ExcelValue model_e77() {
  static ExcelValue result;
  if(variable_set[1221] == 1) { return result;}
  ExcelValue array0[] = {model_e68(),subtract(_common243(),model_e76())};
  result = min(2, array0);
  variable_set[1221] = 1;
  return result;
}

static ExcelValue model_f77() {
  static ExcelValue result;
  if(variable_set[1222] == 1) { return result;}
  ExcelValue array0[] = {model_f68(),subtract(_common244(),model_f76())};
  result = min(2, array0);
  variable_set[1222] = 1;
  return result;
}

static ExcelValue model_g77() {
  static ExcelValue result;
  if(variable_set[1223] == 1) { return result;}
  ExcelValue array0[] = {model_g68(),subtract(_common245(),model_g76())};
  result = min(2, array0);
  variable_set[1223] = 1;
  return result;
}

static ExcelValue model_h77() {
  static ExcelValue result;
  if(variable_set[1224] == 1) { return result;}
  ExcelValue array0[] = {model_h68(),subtract(_common246(),model_h76())};
  result = min(2, array0);
  variable_set[1224] = 1;
  return result;
}

static ExcelValue model_i77() {
  static ExcelValue result;
  if(variable_set[1225] == 1) { return result;}
  ExcelValue array0[] = {model_i68(),subtract(_common247(),model_i76())};
  result = min(2, array0);
  variable_set[1225] = 1;
  return result;
}

static ExcelValue model_j77() {
  static ExcelValue result;
  if(variable_set[1226] == 1) { return result;}
  ExcelValue array0[] = {model_j68(),subtract(_common248(),model_j76())};
  result = min(2, array0);
  variable_set[1226] = 1;
  return result;
}

static ExcelValue model_k77() {
  static ExcelValue result;
  if(variable_set[1227] == 1) { return result;}
  ExcelValue array0[] = {model_k68(),subtract(_common249(),model_k76())};
  result = min(2, array0);
  variable_set[1227] = 1;
  return result;
}

static ExcelValue model_l77() {
  static ExcelValue result;
  if(variable_set[1228] == 1) { return result;}
  ExcelValue array0[] = {model_l68(),subtract(_common250(),model_l76())};
  result = min(2, array0);
  variable_set[1228] = 1;
  return result;
}

static ExcelValue model_m77() {
  static ExcelValue result;
  if(variable_set[1229] == 1) { return result;}
  ExcelValue array0[] = {model_m68(),subtract(_common251(),model_m76())};
  result = min(2, array0);
  variable_set[1229] = 1;
  return result;
}

static ExcelValue model_n77() {
  static ExcelValue result;
  if(variable_set[1230] == 1) { return result;}
  ExcelValue array0[] = {model_n68(),subtract(_common252(),model_n76())};
  result = min(2, array0);
  variable_set[1230] = 1;
  return result;
}

static ExcelValue model_o77() {
  static ExcelValue result;
  if(variable_set[1231] == 1) { return result;}
  ExcelValue array0[] = {model_o68(),subtract(_common253(),model_o76())};
  result = min(2, array0);
  variable_set[1231] = 1;
  return result;
}

static ExcelValue model_p77() {
  static ExcelValue result;
  if(variable_set[1232] == 1) { return result;}
  ExcelValue array0[] = {model_p68(),subtract(_common254(),model_p76())};
  result = min(2, array0);
  variable_set[1232] = 1;
  return result;
}

static ExcelValue model_q77() {
  static ExcelValue result;
  if(variable_set[1233] == 1) { return result;}
  ExcelValue array0[] = {model_q68(),subtract(_common255(),model_q76())};
  result = min(2, array0);
  variable_set[1233] = 1;
  return result;
}

static ExcelValue model_r77() {
  static ExcelValue result;
  if(variable_set[1234] == 1) { return result;}
  ExcelValue array0[] = {model_r68(),subtract(_common256(),model_r76())};
  result = min(2, array0);
  variable_set[1234] = 1;
  return result;
}

static ExcelValue model_s77() {
  static ExcelValue result;
  if(variable_set[1235] == 1) { return result;}
  ExcelValue array0[] = {model_s68(),subtract(_common257(),model_s76())};
  result = min(2, array0);
  variable_set[1235] = 1;
  return result;
}

static ExcelValue model_t77() {
  static ExcelValue result;
  if(variable_set[1236] == 1) { return result;}
  ExcelValue array0[] = {model_t68(),subtract(_common258(),model_t76())};
  result = min(2, array0);
  variable_set[1236] = 1;
  return result;
}

static ExcelValue model_u77() {
  static ExcelValue result;
  if(variable_set[1237] == 1) { return result;}
  ExcelValue array0[] = {model_u68(),subtract(_common259(),model_u76())};
  result = min(2, array0);
  variable_set[1237] = 1;
  return result;
}

static ExcelValue model_v77() {
  static ExcelValue result;
  if(variable_set[1238] == 1) { return result;}
  ExcelValue array0[] = {model_v68(),subtract(_common260(),model_v76())};
  result = min(2, array0);
  variable_set[1238] = 1;
  return result;
}

static ExcelValue model_w77() {
  static ExcelValue result;
  if(variable_set[1239] == 1) { return result;}
  ExcelValue array0[] = {model_w68(),subtract(_common261(),model_w76())};
  result = min(2, array0);
  variable_set[1239] = 1;
  return result;
}

static ExcelValue model_x77() {
  static ExcelValue result;
  if(variable_set[1240] == 1) { return result;}
  ExcelValue array0[] = {model_x68(),subtract(_common262(),model_x76())};
  result = min(2, array0);
  variable_set[1240] = 1;
  return result;
}

static ExcelValue model_y77() {
  static ExcelValue result;
  if(variable_set[1241] == 1) { return result;}
  ExcelValue array0[] = {model_y68(),subtract(_common263(),model_y76())};
  result = min(2, array0);
  variable_set[1241] = 1;
  return result;
}

static ExcelValue model_z77() {
  static ExcelValue result;
  if(variable_set[1242] == 1) { return result;}
  ExcelValue array0[] = {model_z68(),subtract(_common264(),model_z76())};
  result = min(2, array0);
  variable_set[1242] = 1;
  return result;
}

static ExcelValue model_aa77() {
  static ExcelValue result;
  if(variable_set[1243] == 1) { return result;}
  ExcelValue array0[] = {model_aa68(),subtract(_common265(),model_aa76())};
  result = min(2, array0);
  variable_set[1243] = 1;
  return result;
}

static ExcelValue model_ab77() {
  static ExcelValue result;
  if(variable_set[1244] == 1) { return result;}
  ExcelValue array0[] = {model_ab68(),subtract(_common266(),model_ab76())};
  result = min(2, array0);
  variable_set[1244] = 1;
  return result;
}

static ExcelValue model_ac77() {
  static ExcelValue result;
  if(variable_set[1245] == 1) { return result;}
  ExcelValue array0[] = {model_ac68(),subtract(_common267(),model_ac76())};
  result = min(2, array0);
  variable_set[1245] = 1;
  return result;
}

static ExcelValue model_ad77() {
  static ExcelValue result;
  if(variable_set[1246] == 1) { return result;}
  ExcelValue array0[] = {model_ad68(),subtract(_common268(),model_ad76())};
  result = min(2, array0);
  variable_set[1246] = 1;
  return result;
}

static ExcelValue model_ae77() {
  static ExcelValue result;
  if(variable_set[1247] == 1) { return result;}
  ExcelValue array0[] = {model_ae68(),subtract(_common269(),model_ae76())};
  result = min(2, array0);
  variable_set[1247] = 1;
  return result;
}

static ExcelValue model_af77() {
  static ExcelValue result;
  if(variable_set[1248] == 1) { return result;}
  ExcelValue array0[] = {model_af68(),subtract(_common270(),model_af76())};
  result = min(2, array0);
  variable_set[1248] = 1;
  return result;
}

static ExcelValue model_ag77() {
  static ExcelValue result;
  if(variable_set[1249] == 1) { return result;}
  ExcelValue array0[] = {model_ag68(),subtract(_common271(),model_ag76())};
  result = min(2, array0);
  variable_set[1249] = 1;
  return result;
}

static ExcelValue model_ah77() {
  static ExcelValue result;
  if(variable_set[1250] == 1) { return result;}
  ExcelValue array0[] = {model_ah68(),subtract(_common272(),model_ah76())};
  result = min(2, array0);
  variable_set[1250] = 1;
  return result;
}

static ExcelValue model_ai77() {
  static ExcelValue result;
  if(variable_set[1251] == 1) { return result;}
  ExcelValue array0[] = {model_ai68(),subtract(_common273(),model_ai76())};
  result = min(2, array0);
  variable_set[1251] = 1;
  return result;
}

static ExcelValue model_aj77() {
  static ExcelValue result;
  if(variable_set[1252] == 1) { return result;}
  ExcelValue array0[] = {model_aj68(),subtract(_common274(),model_aj76())};
  result = min(2, array0);
  variable_set[1252] = 1;
  return result;
}

static ExcelValue model_ak77() {
  static ExcelValue result;
  if(variable_set[1253] == 1) { return result;}
  ExcelValue array0[] = {model_ak68(),subtract(_common275(),model_ak76())};
  result = min(2, array0);
  variable_set[1253] = 1;
  return result;
}

static ExcelValue model_al77() {
  static ExcelValue result;
  if(variable_set[1254] == 1) { return result;}
  ExcelValue array0[] = {model_al68(),subtract(_common276(),model_al76())};
  result = min(2, array0);
  variable_set[1254] = 1;
  return result;
}

static ExcelValue model_am77() {
  static ExcelValue result;
  if(variable_set[1255] == 1) { return result;}
  ExcelValue array0[] = {model_am68(),subtract(_common277(),model_am76())};
  result = min(2, array0);
  variable_set[1255] = 1;
  return result;
}

static ExcelValue model_an77() {
  static ExcelValue result;
  if(variable_set[1256] == 1) { return result;}
  ExcelValue array0[] = {model_an68(),subtract(_common278(),model_an76())};
  result = min(2, array0);
  variable_set[1256] = 1;
  return result;
}

ExcelValue model_b85() {
  static ExcelValue result;
  if(variable_set[1257] == 1) { return result;}
  result = _common279();
  variable_set[1257] = 1;
  return result;
}

ExcelValue model_c85() {
  static ExcelValue result;
  if(variable_set[1258] == 1) { return result;}
  result = _common283();
  variable_set[1258] = 1;
  return result;
}

ExcelValue model_d85() {
  static ExcelValue result;
  if(variable_set[1259] == 1) { return result;}
  result = _common287();
  variable_set[1259] = 1;
  return result;
}

ExcelValue model_e85() {
  static ExcelValue result;
  if(variable_set[1260] == 1) { return result;}
  result = _common291();
  variable_set[1260] = 1;
  return result;
}

ExcelValue model_f85() {
  static ExcelValue result;
  if(variable_set[1261] == 1) { return result;}
  result = _common295();
  variable_set[1261] = 1;
  return result;
}

ExcelValue model_g85() {
  static ExcelValue result;
  if(variable_set[1262] == 1) { return result;}
  result = _common299();
  variable_set[1262] = 1;
  return result;
}

ExcelValue model_h85() {
  static ExcelValue result;
  if(variable_set[1263] == 1) { return result;}
  result = _common303();
  variable_set[1263] = 1;
  return result;
}

ExcelValue model_i85() {
  static ExcelValue result;
  if(variable_set[1264] == 1) { return result;}
  result = _common307();
  variable_set[1264] = 1;
  return result;
}

ExcelValue model_j85() {
  static ExcelValue result;
  if(variable_set[1265] == 1) { return result;}
  result = _common311();
  variable_set[1265] = 1;
  return result;
}

ExcelValue model_k85() {
  static ExcelValue result;
  if(variable_set[1266] == 1) { return result;}
  result = _common315();
  variable_set[1266] = 1;
  return result;
}

ExcelValue model_l85() {
  static ExcelValue result;
  if(variable_set[1267] == 1) { return result;}
  result = _common319();
  variable_set[1267] = 1;
  return result;
}

ExcelValue model_m85() {
  static ExcelValue result;
  if(variable_set[1268] == 1) { return result;}
  result = _common323();
  variable_set[1268] = 1;
  return result;
}

ExcelValue model_n85() {
  static ExcelValue result;
  if(variable_set[1269] == 1) { return result;}
  result = _common327();
  variable_set[1269] = 1;
  return result;
}

ExcelValue model_o85() {
  static ExcelValue result;
  if(variable_set[1270] == 1) { return result;}
  result = _common331();
  variable_set[1270] = 1;
  return result;
}

ExcelValue model_p85() {
  static ExcelValue result;
  if(variable_set[1271] == 1) { return result;}
  result = _common335();
  variable_set[1271] = 1;
  return result;
}

ExcelValue model_q85() {
  static ExcelValue result;
  if(variable_set[1272] == 1) { return result;}
  result = _common339();
  variable_set[1272] = 1;
  return result;
}

ExcelValue model_r85() {
  static ExcelValue result;
  if(variable_set[1273] == 1) { return result;}
  result = _common343();
  variable_set[1273] = 1;
  return result;
}

ExcelValue model_s85() {
  static ExcelValue result;
  if(variable_set[1274] == 1) { return result;}
  result = _common347();
  variable_set[1274] = 1;
  return result;
}

ExcelValue model_t85() {
  static ExcelValue result;
  if(variable_set[1275] == 1) { return result;}
  result = _common351();
  variable_set[1275] = 1;
  return result;
}

ExcelValue model_u85() {
  static ExcelValue result;
  if(variable_set[1276] == 1) { return result;}
  result = _common355();
  variable_set[1276] = 1;
  return result;
}

ExcelValue model_v85() {
  static ExcelValue result;
  if(variable_set[1277] == 1) { return result;}
  result = _common359();
  variable_set[1277] = 1;
  return result;
}

ExcelValue model_w85() {
  static ExcelValue result;
  if(variable_set[1278] == 1) { return result;}
  result = _common363();
  variable_set[1278] = 1;
  return result;
}

ExcelValue model_x85() {
  static ExcelValue result;
  if(variable_set[1279] == 1) { return result;}
  result = _common367();
  variable_set[1279] = 1;
  return result;
}

ExcelValue model_y85() {
  static ExcelValue result;
  if(variable_set[1280] == 1) { return result;}
  result = _common371();
  variable_set[1280] = 1;
  return result;
}

ExcelValue model_z85() {
  static ExcelValue result;
  if(variable_set[1281] == 1) { return result;}
  result = _common375();
  variable_set[1281] = 1;
  return result;
}

ExcelValue model_aa85() {
  static ExcelValue result;
  if(variable_set[1282] == 1) { return result;}
  result = _common379();
  variable_set[1282] = 1;
  return result;
}

ExcelValue model_ab85() {
  static ExcelValue result;
  if(variable_set[1283] == 1) { return result;}
  result = _common383();
  variable_set[1283] = 1;
  return result;
}

ExcelValue model_ac85() {
  static ExcelValue result;
  if(variable_set[1284] == 1) { return result;}
  result = _common387();
  variable_set[1284] = 1;
  return result;
}

ExcelValue model_ad85() {
  static ExcelValue result;
  if(variable_set[1285] == 1) { return result;}
  result = _common391();
  variable_set[1285] = 1;
  return result;
}

ExcelValue model_ae85() {
  static ExcelValue result;
  if(variable_set[1286] == 1) { return result;}
  result = _common395();
  variable_set[1286] = 1;
  return result;
}

ExcelValue model_af85() {
  static ExcelValue result;
  if(variable_set[1287] == 1) { return result;}
  result = _common399();
  variable_set[1287] = 1;
  return result;
}

ExcelValue model_ag85() {
  static ExcelValue result;
  if(variable_set[1288] == 1) { return result;}
  result = _common403();
  variable_set[1288] = 1;
  return result;
}

ExcelValue model_ah85() {
  static ExcelValue result;
  if(variable_set[1289] == 1) { return result;}
  result = _common407();
  variable_set[1289] = 1;
  return result;
}

ExcelValue model_ai85() {
  static ExcelValue result;
  if(variable_set[1290] == 1) { return result;}
  result = _common411();
  variable_set[1290] = 1;
  return result;
}

ExcelValue model_aj85() {
  static ExcelValue result;
  if(variable_set[1291] == 1) { return result;}
  result = _common415();
  variable_set[1291] = 1;
  return result;
}

ExcelValue model_ak85() {
  static ExcelValue result;
  if(variable_set[1292] == 1) { return result;}
  result = _common419();
  variable_set[1292] = 1;
  return result;
}

ExcelValue model_al85() {
  static ExcelValue result;
  if(variable_set[1293] == 1) { return result;}
  result = _common423();
  variable_set[1293] = 1;
  return result;
}

ExcelValue model_am85() {
  static ExcelValue result;
  if(variable_set[1294] == 1) { return result;}
  result = _common427();
  variable_set[1294] = 1;
  return result;
}

ExcelValue model_an85() {
  static ExcelValue result;
  if(variable_set[1295] == 1) { return result;}
  result = _common431();
  variable_set[1295] = 1;
  return result;
}

static ExcelValue model_k86() {
  static ExcelValue result;
  if(variable_set[1296] == 1) { return result;}
  result = multiply(_common315(),model_k74());
  variable_set[1296] = 1;
  return result;
}

static ExcelValue model_l86() {
  static ExcelValue result;
  if(variable_set[1297] == 1) { return result;}
  result = multiply(_common319(),model_l74());
  variable_set[1297] = 1;
  return result;
}

static ExcelValue model_m86() {
  static ExcelValue result;
  if(variable_set[1298] == 1) { return result;}
  result = multiply(_common323(),model_m74());
  variable_set[1298] = 1;
  return result;
}

static ExcelValue model_n86() {
  static ExcelValue result;
  if(variable_set[1299] == 1) { return result;}
  result = multiply(_common327(),model_n74());
  variable_set[1299] = 1;
  return result;
}

static ExcelValue model_o86() {
  static ExcelValue result;
  if(variable_set[1300] == 1) { return result;}
  result = multiply(_common331(),model_o74());
  variable_set[1300] = 1;
  return result;
}

static ExcelValue model_p86() {
  static ExcelValue result;
  if(variable_set[1301] == 1) { return result;}
  result = multiply(_common335(),model_p74());
  variable_set[1301] = 1;
  return result;
}

static ExcelValue model_q86() {
  static ExcelValue result;
  if(variable_set[1302] == 1) { return result;}
  result = multiply(_common339(),model_q74());
  variable_set[1302] = 1;
  return result;
}

static ExcelValue model_r86() {
  static ExcelValue result;
  if(variable_set[1303] == 1) { return result;}
  result = multiply(_common343(),model_r74());
  variable_set[1303] = 1;
  return result;
}

static ExcelValue model_s86() {
  static ExcelValue result;
  if(variable_set[1304] == 1) { return result;}
  result = multiply(_common347(),model_s74());
  variable_set[1304] = 1;
  return result;
}

static ExcelValue model_t86() {
  static ExcelValue result;
  if(variable_set[1305] == 1) { return result;}
  result = multiply(_common351(),model_t74());
  variable_set[1305] = 1;
  return result;
}

static ExcelValue model_u86() {
  static ExcelValue result;
  if(variable_set[1306] == 1) { return result;}
  result = multiply(_common355(),model_u74());
  variable_set[1306] = 1;
  return result;
}

static ExcelValue model_v86() {
  static ExcelValue result;
  if(variable_set[1307] == 1) { return result;}
  result = multiply(_common359(),model_v74());
  variable_set[1307] = 1;
  return result;
}

static ExcelValue model_w86() {
  static ExcelValue result;
  if(variable_set[1308] == 1) { return result;}
  result = multiply(_common363(),model_w74());
  variable_set[1308] = 1;
  return result;
}

static ExcelValue model_x86() {
  static ExcelValue result;
  if(variable_set[1309] == 1) { return result;}
  result = multiply(_common367(),model_x74());
  variable_set[1309] = 1;
  return result;
}

static ExcelValue model_y86() {
  static ExcelValue result;
  if(variable_set[1310] == 1) { return result;}
  result = multiply(_common371(),model_y74());
  variable_set[1310] = 1;
  return result;
}

static ExcelValue model_z86() {
  static ExcelValue result;
  if(variable_set[1311] == 1) { return result;}
  result = multiply(_common375(),model_z74());
  variable_set[1311] = 1;
  return result;
}

static ExcelValue model_aa86() {
  static ExcelValue result;
  if(variable_set[1312] == 1) { return result;}
  result = multiply(_common379(),model_aa74());
  variable_set[1312] = 1;
  return result;
}

static ExcelValue model_ab86() {
  static ExcelValue result;
  if(variable_set[1313] == 1) { return result;}
  result = multiply(_common383(),model_ab74());
  variable_set[1313] = 1;
  return result;
}

static ExcelValue model_ac86() {
  static ExcelValue result;
  if(variable_set[1314] == 1) { return result;}
  result = multiply(_common387(),model_ac74());
  variable_set[1314] = 1;
  return result;
}

static ExcelValue model_ad86() {
  static ExcelValue result;
  if(variable_set[1315] == 1) { return result;}
  result = multiply(_common391(),model_ad74());
  variable_set[1315] = 1;
  return result;
}

static ExcelValue model_ae86() {
  static ExcelValue result;
  if(variable_set[1316] == 1) { return result;}
  result = multiply(_common395(),model_ae74());
  variable_set[1316] = 1;
  return result;
}

static ExcelValue model_af86() {
  static ExcelValue result;
  if(variable_set[1317] == 1) { return result;}
  result = multiply(_common399(),model_af74());
  variable_set[1317] = 1;
  return result;
}

static ExcelValue model_ag86() {
  static ExcelValue result;
  if(variable_set[1318] == 1) { return result;}
  result = multiply(_common403(),model_ag74());
  variable_set[1318] = 1;
  return result;
}

static ExcelValue model_ah86() {
  static ExcelValue result;
  if(variable_set[1319] == 1) { return result;}
  result = multiply(_common407(),model_ah74());
  variable_set[1319] = 1;
  return result;
}

static ExcelValue model_ai86() {
  static ExcelValue result;
  if(variable_set[1320] == 1) { return result;}
  result = multiply(_common411(),model_ai74());
  variable_set[1320] = 1;
  return result;
}

static ExcelValue model_aj86() {
  static ExcelValue result;
  if(variable_set[1321] == 1) { return result;}
  result = multiply(_common415(),model_aj74());
  variable_set[1321] = 1;
  return result;
}

static ExcelValue model_ak86() {
  static ExcelValue result;
  if(variable_set[1322] == 1) { return result;}
  result = multiply(_common419(),model_ak74());
  variable_set[1322] = 1;
  return result;
}

static ExcelValue model_al86() {
  static ExcelValue result;
  if(variable_set[1323] == 1) { return result;}
  result = multiply(_common423(),model_al74());
  variable_set[1323] = 1;
  return result;
}

static ExcelValue model_am86() {
  static ExcelValue result;
  if(variable_set[1324] == 1) { return result;}
  result = multiply(_common427(),model_am74());
  variable_set[1324] = 1;
  return result;
}

static ExcelValue model_an86() {
  static ExcelValue result;
  if(variable_set[1325] == 1) { return result;}
  result = multiply(_common431(),model_an74());
  variable_set[1325] = 1;
  return result;
}

ExcelValue model_b89() {
  static ExcelValue result;
  if(variable_set[1326] == 1) { return result;}
  ExcelValue array0[] = {subtract(C40,C17),C37};
  result = iferror(divide(subtract(C10,multiply(_common279(),C17)),max(2, array0)),C37);
  variable_set[1326] = 1;
  return result;
}

ExcelValue model_c89() {
  static ExcelValue result;
  if(variable_set[1327] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_c64(),model_c49()),C37};
  result = iferror(divide(subtract(C25,multiply(_common283(),model_c49())),max(2, array0)),C37);
  variable_set[1327] = 1;
  return result;
}

ExcelValue model_d89() {
  static ExcelValue result;
  if(variable_set[1328] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_d64(),model_d49()),C37};
  result = iferror(divide(subtract(model_d48(),multiply(_common287(),model_d49())),max(2, array0)),C37);
  variable_set[1328] = 1;
  return result;
}

ExcelValue model_e89() {
  static ExcelValue result;
  if(variable_set[1329] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_e64(),model_e49()),C37};
  result = iferror(divide(subtract(model_e48(),multiply(_common291(),model_e49())),max(2, array0)),C37);
  variable_set[1329] = 1;
  return result;
}

ExcelValue model_f89() {
  static ExcelValue result;
  if(variable_set[1330] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_f64(),model_f49()),C37};
  result = iferror(divide(subtract(model_f48(),multiply(_common295(),model_f49())),max(2, array0)),C37);
  variable_set[1330] = 1;
  return result;
}

ExcelValue model_g89() {
  static ExcelValue result;
  if(variable_set[1331] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_g64(),model_g49()),C37};
  result = iferror(divide(subtract(model_g48(),multiply(_common299(),model_g49())),max(2, array0)),C37);
  variable_set[1331] = 1;
  return result;
}

ExcelValue model_h89() {
  static ExcelValue result;
  if(variable_set[1332] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_h64(),model_h49()),C37};
  result = iferror(divide(subtract(model_h48(),multiply(_common303(),model_h49())),max(2, array0)),C37);
  variable_set[1332] = 1;
  return result;
}

ExcelValue model_i89() {
  static ExcelValue result;
  if(variable_set[1333] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_i64(),model_i49()),C37};
  result = iferror(divide(subtract(model_i48(),multiply(_common307(),model_i49())),max(2, array0)),C37);
  variable_set[1333] = 1;
  return result;
}

ExcelValue model_j89() {
  static ExcelValue result;
  if(variable_set[1334] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_j64(),model_j49()),C37};
  result = iferror(divide(subtract(model_j48(),multiply(_common311(),model_j49())),max(2, array0)),C37);
  variable_set[1334] = 1;
  return result;
}

ExcelValue model_k89() {
  static ExcelValue result;
  if(variable_set[1335] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_k64(),model_k74()),C37};
  result = iferror(divide(_common19(),max(2, array0)),C37);
  variable_set[1335] = 1;
  return result;
}

ExcelValue model_l89() {
  static ExcelValue result;
  if(variable_set[1336] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_l64(),model_l74()),C37};
  result = iferror(divide(_common20(),max(2, array0)),C37);
  variable_set[1336] = 1;
  return result;
}

ExcelValue model_m89() {
  static ExcelValue result;
  if(variable_set[1337] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_m64(),model_m74()),C37};
  result = iferror(divide(_common21(),max(2, array0)),C37);
  variable_set[1337] = 1;
  return result;
}

ExcelValue model_n89() {
  static ExcelValue result;
  if(variable_set[1338] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_n64(),model_n74()),C37};
  result = iferror(divide(_common22(),max(2, array0)),C37);
  variable_set[1338] = 1;
  return result;
}

ExcelValue model_o89() {
  static ExcelValue result;
  if(variable_set[1339] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_o64(),model_o74()),C37};
  result = iferror(divide(_common23(),max(2, array0)),C37);
  variable_set[1339] = 1;
  return result;
}

ExcelValue model_p89() {
  static ExcelValue result;
  if(variable_set[1340] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_p64(),model_p74()),C37};
  result = iferror(divide(_common24(),max(2, array0)),C37);
  variable_set[1340] = 1;
  return result;
}

ExcelValue model_q89() {
  static ExcelValue result;
  if(variable_set[1341] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_q64(),model_q74()),C37};
  result = iferror(divide(_common25(),max(2, array0)),C37);
  variable_set[1341] = 1;
  return result;
}

ExcelValue model_r89() {
  static ExcelValue result;
  if(variable_set[1342] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_r64(),model_r74()),C37};
  result = iferror(divide(_common26(),max(2, array0)),C37);
  variable_set[1342] = 1;
  return result;
}

ExcelValue model_s89() {
  static ExcelValue result;
  if(variable_set[1343] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_s64(),model_s74()),C37};
  result = iferror(divide(_common27(),max(2, array0)),C37);
  variable_set[1343] = 1;
  return result;
}

ExcelValue model_t89() {
  static ExcelValue result;
  if(variable_set[1344] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_t64(),model_t74()),C37};
  result = iferror(divide(_common4(),max(2, array0)),C37);
  variable_set[1344] = 1;
  return result;
}

ExcelValue model_u89() {
  static ExcelValue result;
  if(variable_set[1345] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_u64(),model_u74()),C37};
  result = iferror(divide(_common28(),max(2, array0)),C37);
  variable_set[1345] = 1;
  return result;
}

ExcelValue model_v89() {
  static ExcelValue result;
  if(variable_set[1346] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_v64(),model_v74()),C37};
  result = iferror(divide(_common29(),max(2, array0)),C37);
  variable_set[1346] = 1;
  return result;
}

ExcelValue model_w89() {
  static ExcelValue result;
  if(variable_set[1347] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_w64(),model_w74()),C37};
  result = iferror(divide(_common30(),max(2, array0)),C37);
  variable_set[1347] = 1;
  return result;
}

ExcelValue model_x89() {
  static ExcelValue result;
  if(variable_set[1348] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_x64(),model_x74()),C37};
  result = iferror(divide(_common31(),max(2, array0)),C37);
  variable_set[1348] = 1;
  return result;
}

ExcelValue model_y89() {
  static ExcelValue result;
  if(variable_set[1349] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_y64(),model_y74()),C37};
  result = iferror(divide(_common32(),max(2, array0)),C37);
  variable_set[1349] = 1;
  return result;
}

ExcelValue model_z89() {
  static ExcelValue result;
  if(variable_set[1350] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_z64(),model_z74()),C37};
  result = iferror(divide(_common33(),max(2, array0)),C37);
  variable_set[1350] = 1;
  return result;
}

ExcelValue model_aa89() {
  static ExcelValue result;
  if(variable_set[1351] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_aa64(),model_aa74()),C37};
  result = iferror(divide(_common34(),max(2, array0)),C37);
  variable_set[1351] = 1;
  return result;
}

ExcelValue model_ab89() {
  static ExcelValue result;
  if(variable_set[1352] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_ab64(),model_ab74()),C37};
  result = iferror(divide(_common35(),max(2, array0)),C37);
  variable_set[1352] = 1;
  return result;
}

ExcelValue model_ac89() {
  static ExcelValue result;
  if(variable_set[1353] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_ac64(),model_ac74()),C37};
  result = iferror(divide(_common36(),max(2, array0)),C37);
  variable_set[1353] = 1;
  return result;
}

ExcelValue model_ad89() {
  static ExcelValue result;
  if(variable_set[1354] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_ad64(),model_ad74()),C37};
  result = iferror(divide(_common37(),max(2, array0)),C37);
  variable_set[1354] = 1;
  return result;
}

ExcelValue model_ae89() {
  static ExcelValue result;
  if(variable_set[1355] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_ae64(),model_ae74()),C37};
  result = iferror(divide(_common38(),max(2, array0)),C37);
  variable_set[1355] = 1;
  return result;
}

ExcelValue model_af89() {
  static ExcelValue result;
  if(variable_set[1356] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_af64(),model_af74()),C37};
  result = iferror(divide(_common39(),max(2, array0)),C37);
  variable_set[1356] = 1;
  return result;
}

ExcelValue model_ag89() {
  static ExcelValue result;
  if(variable_set[1357] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_ag64(),model_ag74()),C37};
  result = iferror(divide(_common40(),max(2, array0)),C37);
  variable_set[1357] = 1;
  return result;
}

ExcelValue model_ah89() {
  static ExcelValue result;
  if(variable_set[1358] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_ah64(),model_ah74()),C37};
  result = iferror(divide(_common41(),max(2, array0)),C37);
  variable_set[1358] = 1;
  return result;
}

ExcelValue model_ai89() {
  static ExcelValue result;
  if(variable_set[1359] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_ai64(),model_ai74()),C37};
  result = iferror(divide(_common42(),max(2, array0)),C37);
  variable_set[1359] = 1;
  return result;
}

ExcelValue model_aj89() {
  static ExcelValue result;
  if(variable_set[1360] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_aj64(),model_aj74()),C37};
  result = iferror(divide(_common43(),max(2, array0)),C37);
  variable_set[1360] = 1;
  return result;
}

ExcelValue model_ak89() {
  static ExcelValue result;
  if(variable_set[1361] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_ak64(),model_ak74()),C37};
  result = iferror(divide(_common44(),max(2, array0)),C37);
  variable_set[1361] = 1;
  return result;
}

ExcelValue model_al89() {
  static ExcelValue result;
  if(variable_set[1362] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_al64(),model_al74()),C37};
  result = iferror(divide(_common45(),max(2, array0)),C37);
  variable_set[1362] = 1;
  return result;
}

ExcelValue model_am89() {
  static ExcelValue result;
  if(variable_set[1363] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_am64(),model_am74()),C37};
  result = iferror(divide(_common46(),max(2, array0)),C37);
  variable_set[1363] = 1;
  return result;
}

ExcelValue model_an89() {
  static ExcelValue result;
  if(variable_set[1364] == 1) { return result;}
  ExcelValue array0[] = {subtract(model_an64(),model_an74()),C37};
  result = iferror(divide(_common10(),max(2, array0)),C37);
  variable_set[1364] = 1;
  return result;
}

ExcelValue model_b56() {
  static ExcelValue result;
  if(variable_set[1365] == 1) { return result;}
  result = BLANK;
  variable_set[1365] = 1;
  return result;
}

ExcelValue model_b54() {
  static ExcelValue result;
  if(variable_set[1366] == 1) { return result;}
  result = BLANK;
  variable_set[1366] = 1;
  return result;
}

// end Model

// Start of named references
ExcelValue average_life_of_low_carbon_generation() {
  static ExcelValue result;
  if(variable_set[1367] == 1) { return result;}
  result = model_b13();
  variable_set[1367] = 1;
  return result;
}

ExcelValue ccs_by_2020() {
  static ExcelValue result;
  if(variable_set[1368] == 1) { return result;}
  result = model_b37();
  variable_set[1368] = 1;
  return result;
}

ExcelValue demand() {
  static ExcelValue result;
  if(variable_set[1369] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b48();
  array0[1] = model_c48();
  array0[2] = model_d48();
  array0[3] = model_e48();
  array0[4] = model_f48();
  array0[5] = model_g48();
  array0[6] = model_h48();
  array0[7] = model_i48();
  array0[8] = model_j48();
  array0[9] = model_k48();
  array0[10] = model_l48();
  array0[11] = model_m48();
  array0[12] = model_n48();
  array0[13] = model_o48();
  array0[14] = model_p48();
  array0[15] = model_q48();
  array0[16] = model_r48();
  array0[17] = model_s48();
  array0[18] = model_t48();
  array0[19] = model_u48();
  array0[20] = model_v48();
  array0[21] = model_w48();
  array0[22] = model_x48();
  array0[23] = model_y48();
  array0[24] = model_z48();
  array0[25] = model_aa48();
  array0[26] = model_ab48();
  array0[27] = model_ac48();
  array0[28] = model_ad48();
  array0[29] = model_ae48();
  array0[30] = model_af48();
  array0[31] = model_ag48();
  array0[32] = model_ah48();
  array0[33] = model_ai48();
  array0[34] = model_aj48();
  array0[35] = model_ak48();
  array0[36] = model_al48();
  array0[37] = model_am48();
  array0[38] = model_an48();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1369] = 1;
  return result;
}

ExcelValue electricity_demand_growth_rate() {
  static ExcelValue result;
  if(variable_set[1370] == 1) { return result;}
  result = model_b32();
  variable_set[1370] = 1;
  return result;
}

ExcelValue electricity_demand_in_2012() {
  static ExcelValue result;
  if(variable_set[1371] == 1) { return result;}
  result = model_b31();
  variable_set[1371] = 1;
  return result;
}

ExcelValue electricity_demand_in_2050() {
  static ExcelValue result;
  if(variable_set[1372] == 1) { return result;}
  result = model_b4();
  variable_set[1372] = 1;
  return result;
}

ExcelValue electricity_emissions_during_cb4() {
  static ExcelValue result;
  if(variable_set[1373] == 1) { return result;}
  result = model_f3();
  variable_set[1373] = 1;
  return result;
}

ExcelValue electrification_start_year() {
  static ExcelValue result;
  if(variable_set[1374] == 1) { return result;}
  result = model_b3();
  variable_set[1374] = 1;
  return result;
}

ExcelValue emissions() {
  static ExcelValue result;
  if(variable_set[1375] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b53();
  array0[1] = model_c53();
  array0[2] = model_d53();
  array0[3] = model_e53();
  array0[4] = model_f53();
  array0[5] = model_g53();
  array0[6] = model_h53();
  array0[7] = model_i53();
  array0[8] = model_j53();
  array0[9] = model_k53();
  array0[10] = model_l53();
  array0[11] = model_m53();
  array0[12] = model_n53();
  array0[13] = model_o53();
  array0[14] = model_p53();
  array0[15] = model_q53();
  array0[16] = model_r53();
  array0[17] = model_s53();
  array0[18] = model_t53();
  array0[19] = model_u53();
  array0[20] = model_v53();
  array0[21] = model_w53();
  array0[22] = model_x53();
  array0[23] = model_y53();
  array0[24] = model_z53();
  array0[25] = model_aa53();
  array0[26] = model_ab53();
  array0[27] = model_ac53();
  array0[28] = model_ad53();
  array0[29] = model_ae53();
  array0[30] = model_af53();
  array0[31] = model_ag53();
  array0[32] = model_ah53();
  array0[33] = model_ai53();
  array0[34] = model_aj53();
  array0[35] = model_ak53();
  array0[36] = model_al53();
  array0[37] = model_am53();
  array0[38] = model_an53();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1375] = 1;
  return result;
}

ExcelValue emissions_factor() {
  static ExcelValue result;
  if(variable_set[1376] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b52();
  array0[1] = model_c52();
  array0[2] = model_d52();
  array0[3] = model_e52();
  array0[4] = model_f52();
  array0[5] = model_g52();
  array0[6] = model_h52();
  array0[7] = model_i52();
  array0[8] = model_j52();
  array0[9] = model_k52();
  array0[10] = model_l52();
  array0[11] = model_m52();
  array0[12] = model_n52();
  array0[13] = model_o52();
  array0[14] = model_p52();
  array0[15] = model_q52();
  array0[16] = model_r52();
  array0[17] = model_s52();
  array0[18] = model_t52();
  array0[19] = model_u52();
  array0[20] = model_v52();
  array0[21] = model_w52();
  array0[22] = model_x52();
  array0[23] = model_y52();
  array0[24] = model_z52();
  array0[25] = model_aa52();
  array0[26] = model_ab52();
  array0[27] = model_ac52();
  array0[28] = model_ad52();
  array0[29] = model_ae52();
  array0[30] = model_af52();
  array0[31] = model_ag52();
  array0[32] = model_ah52();
  array0[33] = model_ai52();
  array0[34] = model_aj52();
  array0[35] = model_ak52();
  array0[36] = model_al52();
  array0[37] = model_am52();
  array0[38] = model_an52();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1376] = 1;
  return result;
}

ExcelValue emissions_factor_2030() {
  static ExcelValue result;
  if(variable_set[1377] == 1) { return result;}
  result = model_f6();
  variable_set[1377] = 1;
  return result;
}

ExcelValue emissions_factor_2050() {
  static ExcelValue result;
  if(variable_set[1378] == 1) { return result;}
  result = model_f7();
  variable_set[1378] = 1;
  return result;
}

ExcelValue high_carbon() {
  static ExcelValue result;
  if(variable_set[1379] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b50();
  array0[1] = model_c50();
  array0[2] = model_d50();
  array0[3] = model_e50();
  array0[4] = model_f50();
  array0[5] = model_g50();
  array0[6] = model_h50();
  array0[7] = model_i50();
  array0[8] = model_j50();
  array0[9] = model_k50();
  array0[10] = model_l50();
  array0[11] = model_m50();
  array0[12] = model_n50();
  array0[13] = model_o50();
  array0[14] = model_p50();
  array0[15] = model_q50();
  array0[16] = model_r50();
  array0[17] = model_s50();
  array0[18] = model_t50();
  array0[19] = model_u50();
  array0[20] = model_v50();
  array0[21] = model_w50();
  array0[22] = model_x50();
  array0[23] = model_y50();
  array0[24] = model_z50();
  array0[25] = model_aa50();
  array0[26] = model_ab50();
  array0[27] = model_ac50();
  array0[28] = model_ad50();
  array0[29] = model_ae50();
  array0[30] = model_af50();
  array0[31] = model_ag50();
  array0[32] = model_ah50();
  array0[33] = model_ai50();
  array0[34] = model_aj50();
  array0[35] = model_ak50();
  array0[36] = model_al50();
  array0[37] = model_am50();
  array0[38] = model_an50();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1379] = 1;
  return result;
}

ExcelValue high_carbon_ef() {
  static ExcelValue result;
  if(variable_set[1380] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b51();
  array0[1] = model_c51();
  array0[2] = model_d51();
  array0[3] = model_e51();
  array0[4] = model_f51();
  array0[5] = model_g51();
  array0[6] = model_h51();
  array0[7] = model_i51();
  array0[8] = model_j51();
  array0[9] = model_k51();
  array0[10] = model_l51();
  array0[11] = model_m51();
  array0[12] = model_n51();
  array0[13] = model_o51();
  array0[14] = model_p51();
  array0[15] = model_q51();
  array0[16] = model_r51();
  array0[17] = model_s51();
  array0[18] = model_t51();
  array0[19] = model_u51();
  array0[20] = model_v51();
  array0[21] = model_w51();
  array0[22] = model_x51();
  array0[23] = model_y51();
  array0[24] = model_z51();
  array0[25] = model_aa51();
  array0[26] = model_ab51();
  array0[27] = model_ac51();
  array0[28] = model_ad51();
  array0[29] = model_ae51();
  array0[30] = model_af51();
  array0[31] = model_ag51();
  array0[32] = model_ah51();
  array0[33] = model_ai51();
  array0[34] = model_aj51();
  array0[35] = model_ak51();
  array0[36] = model_al51();
  array0[37] = model_am51();
  array0[38] = model_an51();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1380] = 1;
  return result;
}

ExcelValue high_carbon_emissions_factor_2012() {
  static ExcelValue result;
  if(variable_set[1381] == 1) { return result;}
  result = model_b40();
  variable_set[1381] = 1;
  return result;
}

ExcelValue high_carbon_emissions_factor_2020() {
  static ExcelValue result;
  if(variable_set[1382] == 1) { return result;}
  result = model_c40();
  variable_set[1382] = 1;
  return result;
}

ExcelValue high_carbon_emissions_factor_2050() {
  static ExcelValue result;
  if(variable_set[1383] == 1) { return result;}
  result = model_d40();
  variable_set[1383] = 1;
  return result;
}

ExcelValue high_carbon_load_factor() {
  static ExcelValue result;
  if(variable_set[1384] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b89();
  array0[1] = model_c89();
  array0[2] = model_d89();
  array0[3] = model_e89();
  array0[4] = model_f89();
  array0[5] = model_g89();
  array0[6] = model_h89();
  array0[7] = model_i89();
  array0[8] = model_j89();
  array0[9] = model_k89();
  array0[10] = model_l89();
  array0[11] = model_m89();
  array0[12] = model_n89();
  array0[13] = model_o89();
  array0[14] = model_p89();
  array0[15] = model_q89();
  array0[16] = model_r89();
  array0[17] = model_s89();
  array0[18] = model_t89();
  array0[19] = model_u89();
  array0[20] = model_v89();
  array0[21] = model_w89();
  array0[22] = model_x89();
  array0[23] = model_y89();
  array0[24] = model_z89();
  array0[25] = model_aa89();
  array0[26] = model_ab89();
  array0[27] = model_ac89();
  array0[28] = model_ad89();
  array0[29] = model_ae89();
  array0[30] = model_af89();
  array0[31] = model_ag89();
  array0[32] = model_ah89();
  array0[33] = model_ai89();
  array0[34] = model_aj89();
  array0[35] = model_ak89();
  array0[36] = model_al89();
  array0[37] = model_am89();
  array0[38] = model_an89();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1384] = 1;
  return result;
}

ExcelValue low_carbon_load_factor() {
  static ExcelValue result;
  if(variable_set[1385] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b85();
  array0[1] = model_c85();
  array0[2] = model_d85();
  array0[3] = model_e85();
  array0[4] = model_f85();
  array0[5] = model_g85();
  array0[6] = model_h85();
  array0[7] = model_i85();
  array0[8] = model_j85();
  array0[9] = model_k85();
  array0[10] = model_l85();
  array0[11] = model_m85();
  array0[12] = model_n85();
  array0[13] = model_o85();
  array0[14] = model_p85();
  array0[15] = model_q85();
  array0[16] = model_r85();
  array0[17] = model_s85();
  array0[18] = model_t85();
  array0[19] = model_u85();
  array0[20] = model_v85();
  array0[21] = model_w85();
  array0[22] = model_x85();
  array0[23] = model_y85();
  array0[24] = model_z85();
  array0[25] = model_aa85();
  array0[26] = model_ab85();
  array0[27] = model_ac85();
  array0[28] = model_ad85();
  array0[29] = model_ae85();
  array0[30] = model_af85();
  array0[31] = model_ag85();
  array0[32] = model_ah85();
  array0[33] = model_ai85();
  array0[34] = model_aj85();
  array0[35] = model_ak85();
  array0[36] = model_al85();
  array0[37] = model_am85();
  array0[38] = model_an85();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1385] = 1;
  return result;
}

ExcelValue maximum_low_c() {
  static ExcelValue result;
  if(variable_set[1386] == 1) { return result;}
  result = model_b12();
  variable_set[1386] = 1;
  return result;
}

ExcelValue maximum_low_carbon_build_rate() {
  static ExcelValue result;
  if(variable_set[1387] == 1) { return result;}
  result = model_b9();
  variable_set[1387] = 1;
  return result;
}

ExcelValue maximum_low_carbon_build_rate_expansion() {
  static ExcelValue result;
  if(variable_set[1388] == 1) { return result;}
  result = model_b11();
  variable_set[1388] = 1;
  return result;
}

ExcelValue maxmean2012() {
  static ExcelValue result;
  if(variable_set[1389] == 1) { return result;}
  result = model_b45();
  variable_set[1389] = 1;
  return result;
}

ExcelValue maxmean2050() {
  static ExcelValue result;
  if(variable_set[1390] == 1) { return result;}
  result = model_c45();
  variable_set[1390] = 1;
  return result;
}

ExcelValue minimum_low_carbon_build_rate() {
  static ExcelValue result;
  if(variable_set[1391] == 1) { return result;}
  result = model_b10();
  variable_set[1391] = 1;
  return result;
}

ExcelValue minmean2012() {
  static ExcelValue result;
  if(variable_set[1392] == 1) { return result;}
  result = model_b44();
  variable_set[1392] = 1;
  return result;
}

ExcelValue minmean2050() {
  static ExcelValue result;
  if(variable_set[1393] == 1) { return result;}
  result = model_c44();
  variable_set[1393] = 1;
  return result;
}

ExcelValue net_increase_in_zero_carbon() {
  static ExcelValue result;
  if(variable_set[1394] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b56();
  array0[1] = model_c56();
  array0[2] = model_d56();
  array0[3] = model_e56();
  array0[4] = model_f56();
  array0[5] = model_g56();
  array0[6] = model_h56();
  array0[7] = model_i56();
  array0[8] = model_j56();
  array0[9] = model_k56();
  array0[10] = model_l56();
  array0[11] = model_m56();
  array0[12] = model_n56();
  array0[13] = model_o56();
  array0[14] = model_p56();
  array0[15] = model_q56();
  array0[16] = model_r56();
  array0[17] = model_s56();
  array0[18] = model_t56();
  array0[19] = model_u56();
  array0[20] = model_v56();
  array0[21] = model_w56();
  array0[22] = model_x56();
  array0[23] = model_y56();
  array0[24] = model_z56();
  array0[25] = model_aa56();
  array0[26] = model_ab56();
  array0[27] = model_ac56();
  array0[28] = model_ad56();
  array0[29] = model_ae56();
  array0[30] = model_af56();
  array0[31] = model_ag56();
  array0[32] = model_ah56();
  array0[33] = model_ai56();
  array0[34] = model_aj56();
  array0[35] = model_ak56();
  array0[36] = model_al56();
  array0[37] = model_am56();
  array0[38] = model_an56();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1394] = 1;
  return result;
}

ExcelValue nuclear_change_2012_2020() {
  static ExcelValue result;
  if(variable_set[1395] == 1) { return result;}
  result = model_b36();
  variable_set[1395] = 1;
  return result;
}

ExcelValue nuclear_in_2012() {
  static ExcelValue result;
  if(variable_set[1396] == 1) { return result;}
  result = model_b35();
  variable_set[1396] = 1;
  return result;
}

ExcelValue renewable_electricity_in_2020() {
  static ExcelValue result;
  if(variable_set[1397] == 1) { return result;}
  result = model_b7();
  variable_set[1397] = 1;
  return result;
}

ExcelValue renewables_in_2012() {
  static ExcelValue result;
  if(variable_set[1398] == 1) { return result;}
  result = model_b34();
  variable_set[1398] = 1;
  return result;
}

ExcelValue year_second_wave_of_building_starts() {
  static ExcelValue result;
  if(variable_set[1399] == 1) { return result;}
  result = model_b8();
  variable_set[1399] = 1;
  return result;
}

ExcelValue zero_carbon() {
  static ExcelValue result;
  if(variable_set[1400] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b49();
  array0[1] = model_c49();
  array0[2] = model_d49();
  array0[3] = model_e49();
  array0[4] = model_f49();
  array0[5] = model_g49();
  array0[6] = model_h49();
  array0[7] = model_i49();
  array0[8] = model_j49();
  array0[9] = model_k49();
  array0[10] = model_l49();
  array0[11] = model_m49();
  array0[12] = model_n49();
  array0[13] = model_o49();
  array0[14] = model_p49();
  array0[15] = model_q49();
  array0[16] = model_r49();
  array0[17] = model_s49();
  array0[18] = model_t49();
  array0[19] = model_u49();
  array0[20] = model_v49();
  array0[21] = model_w49();
  array0[22] = model_x49();
  array0[23] = model_y49();
  array0[24] = model_z49();
  array0[25] = model_aa49();
  array0[26] = model_ab49();
  array0[27] = model_ac49();
  array0[28] = model_ad49();
  array0[29] = model_ae49();
  array0[30] = model_af49();
  array0[31] = model_ag49();
  array0[32] = model_ah49();
  array0[33] = model_ai49();
  array0[34] = model_aj49();
  array0[35] = model_ak49();
  array0[36] = model_al49();
  array0[37] = model_am49();
  array0[38] = model_an49();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1400] = 1;
  return result;
}

ExcelValue zero_carbon_built() {
  static ExcelValue result;
  if(variable_set[1401] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b55();
  array0[1] = model_c55();
  array0[2] = model_d55();
  array0[3] = model_e55();
  array0[4] = model_f55();
  array0[5] = model_g55();
  array0[6] = model_h55();
  array0[7] = model_i55();
  array0[8] = model_j55();
  array0[9] = model_k55();
  array0[10] = model_l55();
  array0[11] = model_m55();
  array0[12] = model_n55();
  array0[13] = model_o55();
  array0[14] = model_p55();
  array0[15] = model_q55();
  array0[16] = model_r55();
  array0[17] = model_s55();
  array0[18] = model_t55();
  array0[19] = model_u55();
  array0[20] = model_v55();
  array0[21] = model_w55();
  array0[22] = model_x55();
  array0[23] = model_y55();
  array0[24] = model_z55();
  array0[25] = model_aa55();
  array0[26] = model_ab55();
  array0[27] = model_ac55();
  array0[28] = model_ad55();
  array0[29] = model_ae55();
  array0[30] = model_af55();
  array0[31] = model_ag55();
  array0[32] = model_ah55();
  array0[33] = model_ai55();
  array0[34] = model_aj55();
  array0[35] = model_ak55();
  array0[36] = model_al55();
  array0[37] = model_am55();
  array0[38] = model_an55();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1401] = 1;
  return result;
}

ExcelValue zero_carbon_decomissioned() {
  static ExcelValue result;
  if(variable_set[1402] == 1) { return result;}
  static ExcelValue array0[39];
  array0[0] = model_b54();
  array0[1] = model_c54();
  array0[2] = model_d54();
  array0[3] = model_e54();
  array0[4] = model_f54();
  array0[5] = model_g54();
  array0[6] = model_h54();
  array0[7] = model_i54();
  array0[8] = model_j54();
  array0[9] = model_k54();
  array0[10] = model_l54();
  array0[11] = model_m54();
  array0[12] = model_n54();
  array0[13] = model_o54();
  array0[14] = model_p54();
  array0[15] = model_q54();
  array0[16] = model_r54();
  array0[17] = model_s54();
  array0[18] = model_t54();
  array0[19] = model_u54();
  array0[20] = model_v54();
  array0[21] = model_w54();
  array0[22] = model_x54();
  array0[23] = model_y54();
  array0[24] = model_z54();
  array0[25] = model_aa54();
  array0[26] = model_ab54();
  array0[27] = model_ac54();
  array0[28] = model_ad54();
  array0[29] = model_ae54();
  array0[30] = model_af54();
  array0[31] = model_ag54();
  array0[32] = model_ah54();
  array0[33] = model_ai54();
  array0[34] = model_aj54();
  array0[35] = model_ak54();
  array0[36] = model_al54();
  array0[37] = model_am54();
  array0[38] = model_an54();
  ExcelValue array0_ev = new_excel_range(array0,1,39);
  result = array0_ev;
  variable_set[1402] = 1;
  return result;
}

void set_maximum_low_carbon_build_rate(ExcelValue newValue) {
  set_model_b9(newValue);
}

void set_year_second_wave_of_building_starts(ExcelValue newValue) {
  set_model_b8(newValue);
}

// End of named references
