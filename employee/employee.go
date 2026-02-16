package employee

// Employee represents a single employee record.
type Employee struct {
	ID         int
	FullName   string
	Position   string
	Department string
}

// GenerateFake returns 10 fake employee records.
func GenerateFake() []Employee {
	return []Employee{
		{ID: 1, FullName: "Atageldi Orazow", Position: "Backend Developer", Department: "IT"},
		{ID: 2, FullName: "Merdan Annayew", Position: "Frontend Developer", Department: "IT"},
		{ID: 3, FullName: "Aynur Saparowa", Position: "Project Manager", Department: "Management"},
		{ID: 4, FullName: "Kerim Muhammedow", Position: "QA Engineer", Department: "IT"},
		{ID: 5, FullName: "Jennet Orazowa", Position: "HR Specialist", Department: "HR"},
		{ID: 6, FullName: "Serdar Berdiýew", Position: "DevOps Engineer", Department: "IT"},
		{ID: 7, FullName: "Ogulgerek Ataýewa", Position: "Accountant", Department: "Finance"},
		{ID: 8, FullName: "Maksat Gurbansähedow", Position: "Designer", Department: "Marketing"},
		{ID: 9, FullName: "Gülälek Baýramowa", Position: "System Administrator", Department: "IT"},
		{ID: 10, FullName: "Döwran Meredow", Position: "Team Lead", Department: "IT"},
	}
}
