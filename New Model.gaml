/**
* Name: NewModel
* Based on the internal empty template. 
* Author: Admin
* Tags: 
*/


model NewModel

/* Insert your model definition here */

global {
    init {
        create people number: 20;
    }
}

species people {
    
    reflex move {
        location <- location + {rnd(3) - 1, rnd(3) - 1};
    }

    aspect default {
        draw circle(2) color: #blue;
    }
}

experiment SimpleExperiment type: gui {
    output {
        display map {
            species people;
        }
    }
}
