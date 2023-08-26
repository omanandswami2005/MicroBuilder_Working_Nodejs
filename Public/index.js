
    const numberOfStudentsDropdown = document.querySelector('select[name="noStd"]');
    const dynamicFieldsContainer = document.getElementById('dynamicFieldsContainer');

    numberOfStudentsDropdown.addEventListener('change', function() {
        const numStudents = parseInt(this.value); // Convert selected value to an integer

        // Clear any existing fields
        dynamicFieldsContainer.innerHTML = '';

        // Generate fields based on selected number
        for (let i = 1; i <= numStudents; i++) {
            dynamicFieldsContainer.innerHTML += `
                <input type="text" name="name${i}" placeholder="Student ${i} Name"><br><br>
                <input type="text" name="enrollment${i}" placeholder="Student ${i} Enrollment"><br><br>
            `;
        }
    });

