# Actuarial-Project-7-Completion-Factors
Explores the calculation of IBNR Reserves using the “completion factor” method.

#### Background
For this project, I am an actuarial analyst working for an insurance company that offers health insurance benefits for employees of various different companies.

This type of insurance is known as “group insurance” because the insurance company provides health insurance benefits to all the employees working at a company (ie a “group” of employees).  With group insurance, the insurance company provides insurance benefits to employees at multiple different companies.  Each company is considered to be it’s own “group” and everyone working at the company (ie everyone in the group) has the same insurance coverage and premium no matter what their age, health status, smoking status, etc.

The insurance policy requires that employee claims are submitted to the insurance company within 6 months of occurring.  Therefore, the maximum amount of lag between when the insurable event (ie a dental appointment or a medication prescription) occurs and when it is reported is 6 months.

#### Task
My task is to create a tool that calculates both the Ultimate Expected Claims and the IBNR Reserve. 

To complete this task, i was given a table of completion factors that another actuary has put together based on a study he completed a few months ago. 

Take note that the completion factors are different depending on a few different factors, such as claim type and employee count. 

The calculator should be set up so it is able to accept any combination of inputs for employee count, claim type, and reported claims by month for January to June.  Then, given your inputs, it will output values for the Ultimate Expected Claims and the IBNR Reserve without having to change any formulas manually. 

For the purposes of this example, the current date is June 30, 2020, the last day of June. 

Employee count is classified into three categories: Small, Medium and Large. 

   Small: 1-100 employees

   Medium: 101-5000 employees

   Large: 5001+ employees
   
Calculations:
**Ultimate Expected Claims For One Month = Reported Claims for the Month x Completion Factor

** January’s Reported Claim will be multiplied by the 6 month completion factor since it is the odds someone has not reported in 6 months. 

IBNR Reserve = Ultimate Expected Claims – Total Claims Reported from January 2020 to June 2020.
   
#### Result

I wrote a VBA macro that takes the inputs (reported claims for Jan-Jun, Employee Count, and Claims Type), pulls the corresponding completion factors, and calculates the IBNR reserve.



   
