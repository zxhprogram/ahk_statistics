package birdwatcher

// TotalBirdCount return the total bird count by summing
// the individual day's counts.
func TotalBirdCount(birdsPerDay []int) int {
	sum := 0
    for index:=range birdsPerDay{
        sum = sum + birdsPerDay[index]
    }
    return sum
}

// BirdsInWeek returns the total bird count by summing
// only the items belonging to the given week.
func BirdsInWeek(birdsPerDay []int, week int) int {
    endIndex:= week * 7 - 1
    startIndex:=week * 7 - 7
    sum:=0
	for index:=range birdsPerDay {
        if index >= startIndex && index <= endIndex {
            sum = sum + birdsPerDay[index]
        }
    }
    return sum
}

// FixBirdCountLog returns the bird counts after correcting
// the bird counts for alternate days.
func FixBirdCountLog(birdsPerDay []int) []int {
	for index:= range birdsPerDay {
        if index % 2 == 0{
            birdsPerDay[index] = birdsPerDay[index] + 1
        }
    }
    return birdsPerDay
}
