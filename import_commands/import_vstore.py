def generate_vstore_command(row):
    """Generate vstore creation command."""
    name = row.get('Vstore', row.get('Name'))
    if not name:
        return None
    
    params = []
    if 'nas_capacity_quota' in row:
        params.append(f"nas_capacity_quota={row['nas_capacity_quota']}")
    if 'description' in row:
        params.append(f"description={row['description']}")
    
    param_str = " " + " ".join(params) if params else ""
    return f"create vstore general name={name}{param_str}"