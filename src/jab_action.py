class JabAction:
    action : str
    name : str
    role : str
    action_type : str
    value : any
    def __init__(self, action, name, role, action_type, value):
        self.action = action
        self.name = name
        self.role = role
        self.action_type = action_type
        self.value = value
    