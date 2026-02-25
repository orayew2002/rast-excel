package domain

import (
	"fmt"
	"math/rand/v2"
	"time"

	"github.com/bxcodec/faker/v4"
)

type Employee struct {
	Id          int
	FullName    string
	TableID     string
	JobPosition string
	Attendance  []string
}

// Mark represents a single attendance legend entry.
// Name is the label shown in the name column; Key is the abbreviation shown
// in the key column (e.g. "B", "W", "IW").
type Mark struct {
	Name string
	Key  string
}

var jobPositions = []string{
	"Software Engineer",
	"Backend Developer",
	"Frontend Developer",
	"DevOps Engineer",
	"QA Engineer",
	"Project Manager",
}

var attendanceSymbols = []string{"W", "8", "P", "W", "A", "L"}

// GenerateEmployees creates n employees with random data.
// Attendance length matches the current month's day count.
func GenerateEmployees(n int) []Employee {
	days := currentMonthDays()
	employees := make([]Employee, n)

	for i := range n {
		employees[i] = Employee{
			Id:          i + 1,
			FullName:    faker.Name(),
			TableID:     fmt.Sprintf("%03d", i+1),
			JobPosition: jobPositions[rand.IntN(len(jobPositions))],
			Attendance:  generateAttendance(days),
		}
	}

	return employees
}

func generateAttendance(days int) []string {
	attendance := make([]string, days)
	for i := range attendance {
		attendance[i] = attendanceSymbols[rand.IntN(len(attendanceSymbols))]
	}
	return attendance
}

func currentMonthDays() int {
	now := time.Now().Local()
	year, month, _ := now.Date()
	first := time.Date(year, month+1, 1, 0, 0, 0, 0, now.Location())
	return first.AddDate(0, 0, -1).Day()
}
