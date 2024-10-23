#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@Time   :2024/10/23 15:28
@Author :lancelot.sheng
@File   :conversation_agent.py
"""
from agent_base import AgentBase


class ConversationAgent(AgentBase):
    """
    对话代理类，负责处理与用户的对话。
    """
    def __init__(self, session_id=None):
        super().__init__(
            name="conversation",
            prompt_file="prompts/formatter.txt",
            session_id=session_id
        )