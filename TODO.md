Refactoring
-----------
* add verbose option for
    - get_date
    - get_weather
    - get_newmail_count


Unit Testing
------------
* get_name
    - return a string
    - string can be:
       - empty 
       - first_name
       - title + last_name
       - formal_address
    - test return until every possible combination is returned

* get_greeting
    - must be a valid string
    - must not be empty
    - must end with a period
    - first part must be Hello,Hi,Good night/morning/afternoon/evening
    - second part must be one of the four examples

* get_signoff
    - must be a valid string
    - must not be empty
    - must end with a period
    - must match one of the four signoffs

* get_date
    - must be format of "It's short_time on long_date"
    - short_time must be format %I:%M%p
    - long_date must be format %A, %B %d
    - must return the current time and date

* get_weather
    - 

* get_newmail_count
    - incorrect login

* get_tasks
    - must return range [0,inf]
    - must be a number
