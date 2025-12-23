// Package weather is a test file.
package weather

var (
    // CurrentCondition is a string type.
	CurrentCondition string
    // CurrentLocation is a string type.
	CurrentLocation  string
)
// Forecast is a function for concat city and contion to a full string.
func Forecast(city, condition string) string {
	CurrentLocation, CurrentCondition = city, condition
	return CurrentLocation + " - current weather condition: " + CurrentCondition
}
