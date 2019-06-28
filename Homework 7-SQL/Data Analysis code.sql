--1. List the following details of each employee: employee number, last name, first name, gender, and salary.
Select employees.emp_no, employees.last_name, employees.first_name,employees.gender,salaries.salary
from employees
Join salaries on employees.emp_no = salaries.emp_no;

--2. List employees who were hired in 1986.
select hire_date
from employees
where EXTRACT(YEAR FROM hire_date)= '1986';

--3. List the manager of each department with the following information: department number, department name, the manager's employee number, last name, first name, and start and end employment dates.
select dept_manager.dept_no,departments.dept_name,dept_manager.emp_no,employees.last_name,employees.first_name,dept_manager.from_date,dept_manager.to_date
from dept_manager
join departments as departments
on departments.dept_no = dept_manager.dept_no
join employees as employees
on employees.emp_no = dept_manager.emp_no;

--4. List the department of each employee with the following information: employee number, last name, first name, and department name.
select employees.emp_no,employees.last_name,employees.first_name,departments.dept_name
from employees
join dept_emp as dept_emp
on dept_emp.emp_no=employees.emp_no
join departments as departments
on departments.dept_no = dept_emp.dept_no;

--5. List all employees whose first name is "Hercules" and last names begin with "B."
select first_name,last_name
from employees
where first_name = 'Hercules' and last_name like ('B%') 

--6. List all employees in the Sales department, including their employee number, last name, first name, and department name.
select employees.emp_no,employees.last_name,employees.first_name,departments.dept_name
from employees
join dept_emp as dept_emp
on dept_emp.emp_no=employees.emp_no
join departments as departments
on departments.dept_no = dept_emp.dept_no
where departments.dept_name = 'Sales';

--7. List all employees in the Sales and Development departments, including their employee number, last name, first name, and department name.
select employees.emp_no,employees.last_name,employees.first_name,departments.dept_name
from employees
join dept_emp as dept_emp
on dept_emp.emp_no=employees.emp_no
join departments as departments
on departments.dept_no = dept_emp.dept_no
where departments.dept_name = 'Sales'or departments.dept_name = 'Development';

--8. In descending order, list the frequency count of employee last names, i.e., how many employees share each last name.
SELECT employees.last_name,count(employees.last_name)
FROM employees
GROUP BY last_name 
order by last_name desc; 