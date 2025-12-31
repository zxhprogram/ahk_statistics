package lasagna

// TODO: define the 'PreparationTime()' function
func PreparationTime(layers []string, time int) int{
    if time == 0 {
        time = 2
    }
    return len(layers) * time
}
// TODO: define the 'Quantities()' function
func Quantities(layers []string) (int,float64){
    sumOfNoodles:=0
    sumOfSauce:=0.0
    for index:= range layers {
        if layers[index] == "sauce" {
            sumOfSauce = sumOfSauce + 0.2
        }
        if layers[index] == "noodles" {
            sumOfNoodles = sumOfNoodles + 50
        }
    }
    return sumOfNoodles,sumOfSauce
}

// TODO: define the 'AddSecretIngredient()' function
func AddSecretIngredient(f []string, m []string) {
	m[len(m)-1] = f[len(f)-1]
}


// TODO: define the 'ScaleRecipe()' function
func ScaleRecipe(quantities []float64,num int) []float64{
    var r []float64
    s:=float64(num) / float64(2)
    for index:=range quantities {
        r1 := quantities[index] * s
        r = append(r,r1)
    }
    return r
}

// Your first steps could be to read through the tasks, and create
// these functions with their correct parameter lists and return types.
// The function body only needs to contain `panic("")`.
//
// This will make the tests compile, but they will fail.
// You can then implement the function logic one by one and see
// an increasing number of tests passing as you implement more
// functionality.
