def map_project(assignment, client):
    if client == 'John Smith':
        return 'Project1'
    if client == 'Client1' and assignment == 'Subproject 1':
        return 'Some Fancy Project'
    raise ValueError(f"Unmapped {assignment}, {client}")

