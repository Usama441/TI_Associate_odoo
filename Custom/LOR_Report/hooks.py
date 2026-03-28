def post_init_hook(env):
    env['res.company']._sync_lor_template_defaults_to_latest_source()
